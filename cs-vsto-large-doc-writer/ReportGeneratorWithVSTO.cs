using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Drawing;
using LargeDocWriter.Elements;

namespace LargeDocWriter
{
    public class ReportGeneratorWithVSTO : ReportGenerator
    {
        protected string mContentFolderName = null;
        protected string mContentFolderPath = null;

        private int mSectionIndex = 0;
        private int mSubSectionIndex = 0;
        private int mSubSubSectionIndex = 0;
        private int mSubSubSubSectionIndex = 0;

        private bool mGenerateSectionNumber = true;


        public ReportGeneratorWithVSTO(ReportModel model)
            : base(model)
        {

        }

        public override void GenerateReport(string filepath, string content_folder_path)
        {
            mContentFolderPath = content_folder_path;
            mContentFolderName = Path.GetFileName(content_folder_path);

            object missing = Type.Missing;
            object notTrue = false;

            Word.Application myApp = null;
            Word.Document myDoc = null;

            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */ 

            try
            {
                myApp = new Word.Application();
                myApp.Visible = false;

                myDoc = myApp.Documents.Add();

                //Method 1 to prevent the stupid error dialog ""
                myDoc.GrammarChecked = true;
                myDoc.SpellingChecked = true;

                //Method 2 to prevent the stupid error dialog ""
                myDoc.ShowGrammaticalErrors = false;
                myDoc.ShowRevisions = false;
                myDoc.ShowSpellingErrors = false;

                List<ReportSection> sections = mModel.Sections;
                int section_count = sections.Count;
                for (int sindex = 0; sindex < section_count; ++sindex)
                {
                    ReportSection section = sections[sindex];
                    GenerateSection(myApp, myDoc, ref missing, section);
                }

                NotifyMessage(">> Update Header and Footer");

                if (!string.IsNullOrEmpty(mModel.LeftHeaderImagePath) && File.Exists(mModel.LeftHeaderImagePath))
                {
                    LibHelperWithVSTO.AddHeaderImageLeft(myApp, myDoc, ref missing, mModel.LeftHeaderContent);
                }
                if (!string.IsNullOrEmpty(mModel.LeftHeaderContent))
                {
                    LibHelperWithVSTO.AddHeaderLeft(myApp, myDoc, ref missing, mModel.LeftHeaderContent);
                }

                LibHelperWithVSTO.AddFooterPageNumberRight(myApp, myDoc, ref missing);

                NotifyMessage(">> Update Cross References");
                GenerateCrossReferences(myApp, myDoc);

                object fileFormat = Word.WdSaveFormat.wdFormatXMLDocument;

                //Method 1 to prevent the stupid error dialog ""
                myDoc.GrammarChecked = true;
                myDoc.SpellingChecked = true;

                myDoc.SaveAs(filepath, ref fileFormat, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing);

                ((Word._Document)myDoc).Close(ref missing, ref missing, ref missing);

                ((Word._Application)myApp).Quit(ref notTrue, ref missing, ref missing);

                myDoc = null;
                myApp = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine("GenerateReport throws the error: {0}", ex.ToString());
                ReportError(ex.ToString());
            }

            // Clean up the unmanaged Word COM resources by forcing a garbage 
            // collection as soon as the calling function is off the stack (at 
            // which point these objects are no longer rooted).

            GC.Collect();
            GC.WaitForPendingFinalizers();
            // GC needs to be called twice in order to get the Finalizers called 
            // - the first time in, it simply makes a list of what is to be 
            // finalized, the second time in, it actually is finalizing. Only 
            // then will the object do its automatic ReleaseComObject.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            
        }

        private void GenerateCrossReferences(Word.Application myApp, Word.Document myDoc)
        {
            List<string> parts;
            List<int> part_types;
            Dictionary<string, string> tags;
            Dictionary<string, string> names;

            int paragraph_index = 0;
            int paragraph_count = myDoc.Content.Paragraphs.Count;

            DateTime ticked_time = DateTime.Now;
            DateTime interval_time = ticked_time;
            DateTime start_time = ticked_time;

            foreach(Word.Paragraph p in myDoc.Content.Paragraphs)
            {
                ticked_time = DateTime.Now;
                TimeSpan ts = ticked_time - interval_time;
                paragraph_index++;
                if (ts.TotalMilliseconds > 1000)
                {
                    interval_time = ticked_time;
                    int progress_percentage = paragraph_index * 100 / paragraph_count;

                    TimeSpan duration = ticked_time - start_time;
                    float duration_in_minutes = (float)duration.TotalMinutes;
                    float remaining_duration_in_minutes = -1;
                    if (progress_percentage > 0)
                    {
                        remaining_duration_in_minutes = duration_in_minutes * (100 - progress_percentage) / progress_percentage;
                    }
                    NotifyTaskProgressChanged(string.Format("Update Cross Reference in Paragraph #{0} ({1}%) (Dur: {2:0.0} min Remain Dur: {3:0.0} min)", paragraph_index, progress_percentage, duration_in_minutes, remaining_duration_in_minutes), progress_percentage);
                }

                string text = p.Range.Text;

                string pattern = @"\[Anchor\|Tag:([A-Za-z0-9\-\. ]+)\|Name:([A-Za-z0-9\-\. ]+)\]";

                parts = TextHelper.SplitText(text, pattern, out part_types, out tags, out names);

                if (parts.Count == 0) continue;

                //p.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                StringBuilder sb = new StringBuilder();

                for(int i=0; i < parts.Count; ++i)
                {
                    if(TextHelper.IsRegularText(part_types[i]))
                    {
                        //p.Range.InsertAfter(parts[i]);
                        //p.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        sb.Append(parts[i]);
                    }
                    else if (TextHelper.IsCrossReference(part_types[i]))
                    {
                        bool cf_found = false;
                        if (tags.ContainsKey(parts[i]))
                        {
                            string tag2 = tags[parts[i]];
                            if (mFigureCrossReferenceItems.ContainsKey(tag2))
                            {
                                cf_found = true;
                                //p.Range.InsertCrossReference("Figure", Word.WdReferenceKind.wdOnlyLabelAndNumber, mFigureCrossReferenceItems[tag2], true);
                                sb.AppendFormat("Figure {0}", mFigureCrossReferenceItems[tag2]);
                            }
                            else if (mTableCrossReferenceItems.ContainsKey(tag2))
                            {
                                cf_found = true;
                                sb.AppendFormat("Table {0}", mTableCrossReferenceItems[tag2]);
                            }
                        }

                        if(!cf_found)
                        {
                            Console.WriteLine(parts[i]);
                            //p.Range.InsertAfter(parts[i]);
                            //p.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                            if (names.ContainsKey(parts[i]))
                            {
                                sb.Append(names[parts[i]]);
                            }
                            else
                            {
                                sb.Append(parts[i]);
                            }
                            
                        }
                    }
                }

                p.Range.Text = sb.ToString();
                p.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
            }
            NotifyTaskProgressChanged("Cross Reference Completed", 0);
        }

        private void GenerateSection(Word.Application myApp, Word.Document myDoc, ref object missing, ReportSection section)
        {

            NotifyMessage(string.Format(">> {0}", section.Header.Title));

            AddHeading(myApp, myDoc, ref missing, section.Header);

            List<object> parts = section.Parts;
            int part_count = parts.Count;
            for (int pindex = 0; pindex < part_count; ++pindex)
            {
                object part = parts[pindex];
                if (part is ReportSection)
                {
                    ReportSection subsection = part as ReportSection;
                    GenerateSection(myApp, myDoc, ref missing, subsection);
                }
                else if (part is ReportParagraph)
                {
                    ReportParagraph p=part as ReportParagraph;
                    GenerateParagraph(myApp, myDoc, ref missing, p);
                }
                else if (part is ReportFigure)
                {
                    ReportFigure figure = part as ReportFigure;
                    GenerateFigure(myApp, myDoc, ref missing, figure);
                }
                else if (part is ReportTable)
                {
                    ReportTable table = part as ReportTable;
                    GenerateTable(myApp, myDoc, ref missing, table);
                }
            }
        }

        private void GenerateParagraph(Word.Application myApp, Word.Document myDoc, ref object missing, ReportParagraph p)
        {
            string p_content = p.Content;
            //p_content = Regex.Replace(p_content, @"\[Anchor\|Tag:([A-Za-z0-9\-\. ]+)\|Name:([A-Za-z0-9\-\. ]+)\]", (match) =>
            //{
            //    string tag2 = match.Groups[1].Value;
            //    string name2 = match.Groups[2].Value;

            //    return string.Format("{0}", name2);
            //});

            LibHelperWithVSTO.AddParagraph(myApp, myDoc, ref missing, p_content);
        }

        private Dictionary<string, int> mFigureCrossReferenceItems = new Dictionary<string, int>();
        private Dictionary<string, int> mTableCrossReferenceItems = new Dictionary<string, int>();

        private void GenerateFigure(Word.Application myApp, Word.Document myDoc, ref object missing, ReportFigure figure)
        {
            string figure_filename = figure.FileName;
            string figure_filepath = GetImageFullPath(figure_filename);

            if (figure is ReportFigure_PieChart)
            {
                ReportFigure_PieChart chart = figure as ReportFigure_PieChart;
                //LibHelperWithChartDirector.WriteToPieChart(chart.Content, chart.Width, chart.Height, figure_filepath);
                //GenerateFigure(myApp, myDoc, ref missing, figure, figure_filepath);
                GenerateFigure_PieChart(myApp, myDoc, ref missing, chart);
            }
            else if (figure is ReportFigure_ColumnChart)
            {
                ReportFigure_ColumnChart chart = figure as ReportFigure_ColumnChart;
                //LibHelperWithChartDirector.WriteToBarChart(chart.Content, chart.XLabel, chart.YLabel, chart.Width, chart.Height, figure_filepath);
                //GenerateFigure(myApp, myDoc, ref missing, figure, figure_filepath);
                GenerateFigure_ColumnChart(myApp, myDoc, ref missing, chart);
            }
            else if (figure is ReportFigure_BarChart)
            {
                ReportFigure_BarChart chart = figure as ReportFigure_BarChart;
                GenerateFigure_BarChart(myApp, myDoc, ref missing, chart);
            }
            else
            {
                GenerateFigure(myApp, myDoc, ref missing, figure, figure_filepath);
            }
        }

        private void GenerateFigure_ColumnChart(Word.Application myApp, Word.Document myDoc, ref object missing, ReportFigure_ColumnChart figure)
        {
            string figure_caption = figure.Caption;

            figure_caption = Regex.Replace(figure_caption, @"\[Anchor\|Tag:([A-Za-z0-9\-\. ]+)\|Name:([A-Za-z0-9\-\. ]+)\]", (match) =>
            {
                string tag2 = match.Groups[1].Value;
                string name2 = match.Groups[2].Value;

                return string.Format("{0}", name2);
            });

            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

            object oRng = myDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            Word.Paragraph oPara = myDoc.Content.Paragraphs.Add(ref oRng);
            oPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            oPara.Range.InsertParagraphAfter();

            //string img_caption = "image caption";

            Word.InlineShape wdShape = myDoc.Content.InlineShapes.AddChart(Microsoft.Office.Core.XlChartType.xl3DColumn, oPara.Range);
            Word.Chart wdChart = wdShape.Chart;

            Word.ChartData chartData = wdChart.ChartData;

            Excel.Workbook dataWorkbook = (Excel.Workbook)chartData.Workbook;
            Excel.Worksheet dataSheet = (Excel.Worksheet)dataWorkbook.Worksheets[1];

            dataWorkbook.Application.Visible = false;

            Dictionary<string, float> data_table = figure.Content;
            int row_count = data_table.Count;

            Excel.Range tRange = dataSheet.Cells.get_Range("A1", string.Format("B{0}", row_count+1));
            Excel.ListObject tbl1 = dataSheet.ListObjects["Table1"];
            tbl1.Resize(tRange);

            List<string> data_keys = data_table.Keys.ToList();

            ((Excel.Range)dataSheet.Cells.get_Range("B1", missing)).FormulaR1C1 = "Histogram";
            for(int i=0; i < row_count; ++i)
            {
                string data_key = data_keys[i];
                ((Excel.Range)dataSheet.Cells.get_Range(string.Format("A{0}", i+2), missing)).FormulaR1C1 = data_key;
                ((Excel.Range)dataSheet.Cells.get_Range(string.Format("B{0}", i+2), missing)).FormulaR1C1 = data_table[data_key];
            }

            if (wdChart.HasTitle)
            {
                wdChart.ChartTitle.Font.Italic = true;
                wdChart.ChartTitle.Font.Size = 18;
                wdChart.ChartTitle.Font.Color = Color.Black.ToArgb();
                wdChart.ChartTitle.Text = figure.Caption;
                wdChart.ChartTitle.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;

                wdChart.ChartTitle.Format.Line.ForeColor.RGB = Color.Black.ToArgb();
            }

            wdChart.RightAngleAxes = true;

            wdChart.ApplyDataLabels(Word.XlDataLabelsType.xlDataLabelsShowLabel, missing, missing, missing, missing, missing, missing, missing, missing, missing);

            wdChart.Legend.Delete();

            object oLabel = Word.WdCaptionLabelID.wdCaptionFigure;
            object oTitle = string.Format(" {0}", figure.Caption);
            object position = Word.WdCaptionPosition.wdCaptionPositionBelow;
            wdShape.Range.InsertCaption(ref oLabel, ref oTitle, ref missing, ref position);

            object arr_r = myDoc.GetCrossReferenceItems("Figure");
            Array arr = ((Array)(arr_r));

            mFigureCrossReferenceItems[figure.Tag] = arr.GetUpperBound(0);
        }

        private void GenerateFigure_BarChart(Word.Application myApp, Word.Document myDoc, ref object missing, ReportFigure_BarChart figure)
        {
            string figure_caption = figure.Caption;

            figure_caption = Regex.Replace(figure_caption, @"\[Anchor\|Tag:([A-Za-z0-9\-\. ]+)\|Name:([A-Za-z0-9\-\. ]+)\]", (match) =>
            {
                string tag2 = match.Groups[1].Value;
                string name2 = match.Groups[2].Value;

                return string.Format("{0}", name2);
            });

            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

            object oRng = myDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            Word.Paragraph oPara = myDoc.Content.Paragraphs.Add(ref oRng);
            oPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            oPara.Range.InsertParagraphAfter();

            //string img_caption = "image caption";

            Word.InlineShape wdShape = myDoc.Content.InlineShapes.AddChart(Microsoft.Office.Core.XlChartType.xl3DBarClustered, oPara.Range);
            Word.Chart wdChart = wdShape.Chart;

            Word.ChartData chartData = wdChart.ChartData;

            Excel.Workbook dataWorkbook = (Excel.Workbook)chartData.Workbook;
            Excel.Worksheet dataSheet = (Excel.Worksheet)dataWorkbook.Worksheets[1];

            dataWorkbook.Application.Visible = false;

            Dictionary<string, float> data_table = figure.Content;
            int row_count = data_table.Count;

            Excel.Range tRange = dataSheet.Cells.get_Range("A1", string.Format("B{0}", row_count + 1));
            Excel.ListObject tbl1 = dataSheet.ListObjects["Table1"];
            tbl1.Resize(tRange);

            List<string> data_keys = data_table.Keys.ToList();

            ((Excel.Range)dataSheet.Cells.get_Range("B1", missing)).FormulaR1C1 = "Histogram";
            for (int i = 0; i < row_count; ++i)
            {
                string data_key = data_keys[i];
                ((Excel.Range)dataSheet.Cells.get_Range(string.Format("A{0}", i + 2), missing)).FormulaR1C1 = data_key;
                ((Excel.Range)dataSheet.Cells.get_Range(string.Format("B{0}", i + 2), missing)).FormulaR1C1 = data_table[data_key];
            }

            wdChart.ChartTitle.Font.Italic = true;
            wdChart.ChartTitle.Font.Size = 18;
            wdChart.ChartTitle.Font.Color = Color.Black.ToArgb();
            wdChart.ChartTitle.Text = figure.Caption;
            wdChart.ChartTitle.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;

            wdChart.ChartTitle.Format.Line.ForeColor.RGB = Color.Black.ToArgb();

            wdChart.RightAngleAxes = true;

            wdChart.ApplyDataLabels(Word.XlDataLabelsType.xlDataLabelsShowLabel, missing, missing, missing, missing, missing, missing, missing, missing, missing);

            wdChart.Legend.Delete();

            object oLabel = Word.WdCaptionLabelID.wdCaptionFigure;
            object oTitle = string.Format(" {0}", figure.Caption);
            object position = Word.WdCaptionPosition.wdCaptionPositionBelow;
            wdShape.Range.InsertCaption(ref oLabel, ref oTitle, ref missing, ref position);

            object arr_r = myDoc.GetCrossReferenceItems("Figure");
            Array arr = ((Array)(arr_r));

            mFigureCrossReferenceItems[figure.Tag] = arr.GetUpperBound(0);
        }

        private void GenerateFigure_PieChart(Word.Application myApp, Word.Document myDoc, ref object missing, ReportFigure_PieChart figure)
        {
            string figure_caption = figure.Caption;

            figure_caption = Regex.Replace(figure_caption, @"\[Anchor\|Tag:([A-Za-z0-9\-\. ]+)\|Name:([A-Za-z0-9\-\. ]+)\]", (match) =>
            {
                string tag2 = match.Groups[1].Value;
                string name2 = match.Groups[2].Value;

                return string.Format("{0}", name2);
            });

            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

            object oRng = myDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            Word.Paragraph oPara = myDoc.Content.Paragraphs.Add(ref oRng);
            oPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            oPara.Range.InsertParagraphAfter();

            //string img_caption = "image caption";

            Word.InlineShape wdShape = myDoc.Content.InlineShapes.AddChart(Microsoft.Office.Core.XlChartType.xl3DPieExploded, oPara.Range);
            Word.Chart wdChart = wdShape.Chart;

            Word.ChartData chartData = wdChart.ChartData;

            Excel.Workbook dataWorkbook = (Excel.Workbook)chartData.Workbook;
            Excel.Worksheet dataSheet = (Excel.Worksheet)dataWorkbook.Worksheets[1];

            dataWorkbook.Application.Visible = false;

            Dictionary<string, float> data_table = figure.Content;
            int row_count = data_table.Count;

            Excel.Range tRange = dataSheet.Cells.get_Range("A1", string.Format("B{0}", row_count + 1));
            Excel.ListObject tbl1 = dataSheet.ListObjects["Table1"];
            tbl1.Resize(tRange);

            List<string> data_keys = data_table.Keys.ToList();

            for (int i = 0; i < row_count; ++i)
            {
                string data_key = data_keys[i];
                ((Excel.Range)dataSheet.Cells.get_Range(string.Format("A{0}", i + 2), missing)).FormulaR1C1 = data_key;
                ((Excel.Range)dataSheet.Cells.get_Range(string.Format("B{0}", i + 2), missing)).FormulaR1C1 = data_table[data_key];
            }

            wdChart.ChartTitle.Font.Italic = true;
            wdChart.ChartTitle.Font.Size = 18;
            wdChart.ChartTitle.Font.Color = Color.Black.ToArgb();
            wdChart.ChartTitle.Text = figure.Caption;
            wdChart.ChartTitle.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;

            wdChart.ChartTitle.Format.Line.ForeColor.RGB = Color.Black.ToArgb();

            wdChart.ApplyDataLabels(Word.XlDataLabelsType.xlDataLabelsShowLabel, missing, missing, missing, missing, missing, missing, missing, missing, missing);

            object oLabel = Word.WdCaptionLabelID.wdCaptionFigure;
            object oTitle = string.Format(" {0}", figure.Caption);
            object position = Word.WdCaptionPosition.wdCaptionPositionBelow;
            wdShape.Range.InsertCaption(ref oLabel, ref oTitle, ref missing, ref position);

            object arr_r = myDoc.GetCrossReferenceItems("Figure");
            Array arr = ((Array)(arr_r));

            mFigureCrossReferenceItems[figure.Tag] = arr.GetUpperBound(0);
        }

        private void GenerateFigure(Word.Application myApp, Word.Document myDoc, ref object missing, ReportFigure figure, string figure_filepath)
        {
            string figure_caption = figure.Caption;

            figure_caption = Regex.Replace(figure_caption, @"\[Anchor\|Tag:([A-Za-z0-9\-\. ]+)\|Name:([A-Za-z0-9\-\. ]+)\]", (match) =>
            {
                string tag2 = match.Groups[1].Value;
                string name2 = match.Groups[2].Value;

                return string.Format("{0}", name2);
            });

            LibHelperWithVSTO.AddImageFromFile(myApp, myDoc, ref missing, figure_filepath, figure_caption);

            object arr_r = myDoc.GetCrossReferenceItems("Figure");
            Array arr = ((Array)(arr_r));

            mFigureCrossReferenceItems[figure.Tag] = arr.GetUpperBound(0);
        }

        private void GenerateTable(Word.Application myApp, Word.Document myDoc, ref object missing, ReportTable table)
        {
            DateTime ticked_time = DateTime.Now;
            DateTime interval_time = DateTime.Now;

            LibHelperWithVSTO.AddTableFromDataTable(myApp, myDoc, ref missing, table.Content, table.Caption, (s, e) =>
                {
                    ticked_time = DateTime.Now;
                    TimeSpan ts = ticked_time - interval_time;
                    if (ts.TotalMilliseconds > 1000)
                    {
                        interval_time = ticked_time;
                        NotifyTaskProgressChanged(e.UserState as string, e.ProgressPercentage);
                    }
                });

            object arr_r = myDoc.GetCrossReferenceItems("Table");
            Array arr = ((Array)(arr_r));

            mTableCrossReferenceItems[table.Tag] = arr.GetUpperBound(0);
        }

        private void AddHeading(Word.Application myApp, Word.Document myDoc, ref object missing, ReportHeader header)
        {
            int level = header.Level;
            if (level == 1)
            {
                AddHeading1(myApp, myDoc, ref missing, header.Title);
            }
            else if (level == 2)
            {
                AddHeading2(myApp, myDoc, ref missing, header.Title);
            }
            else if (level == 3)
            {
                AddHeading3(myApp, myDoc, ref missing, header.Title);
            }
            else if (level == 4)
            {
                AddHeading4(myApp, myDoc, ref missing, header.Title);
            }
        }

        private string GetImageFullPath(string image_filename)
        {
            return Path.Combine(mContentFolderPath, image_filename);
        }

        private void AddHeading1(Word.Application myApp, Word.Document myDoc, ref object missing, string content)
        {
            //Console.WriteLine("AddHeading1: {0}", content);
            mSectionIndex++;
            mSubSectionIndex = 0;
            mSubSubSectionIndex = 0;
            mSubSubSubSectionIndex = 0;
            LibHelperWithVSTO.AddHeading1(myApp, myDoc, ref missing, content, mGenerateSectionNumber, mSectionIndex);
        }

        private void AddHeading2(Word.Application myApp, Word.Document myDoc, ref object missing, string content)
        {
            mSubSectionIndex++;
            //Console.WriteLine("AddHeading2: {0}", content);
            mSubSubSectionIndex = 0;
            mSubSubSubSectionIndex = 0;
            LibHelperWithVSTO.AddHeading2(myApp, myDoc, ref missing, content, mGenerateSectionNumber, mSectionIndex, mSubSectionIndex);
        }

        private void AddHeading3(Word.Application myApp, Word.Document myDoc, ref object missing, string content)
        {
            mSubSubSectionIndex++;
            mSubSubSubSectionIndex = 0;
            LibHelperWithVSTO.AddHeading3(myApp, myDoc, ref missing, content, mGenerateSectionNumber, mSectionIndex, mSubSectionIndex, mSubSubSectionIndex);
        }

        private void AddHeading4(Word.Application myApp, Word.Document myDoc, ref object missing, string content)
        {
            mSubSubSubSectionIndex++;
            if (mGenerateSectionNumber)
            {
                LibHelperWithVSTO.AddHeading3(myApp, myDoc, ref missing, content, mGenerateSectionNumber, mSectionIndex, mSubSectionIndex, mSubSubSectionIndex, mSubSubSubSectionIndex);
            }
            else
            {
                LibHelperWithVSTO.AddHeading4(myApp, myDoc, ref missing, content, mGenerateSectionNumber, mSectionIndex, mSubSectionIndex, mSubSubSectionIndex, mSubSubSubSectionIndex);
            }
        }
    }
}
