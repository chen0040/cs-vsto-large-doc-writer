using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using System.ComponentModel; 

namespace LargeDocWriter
{
    public class LibHelperWithVSTO
    {
        private static object oEndOfDoc = @"\endofdoc";    // A predefined bookmark 
        private static int mMaxTableRowCount = 32760 - 2;

        public static int MaxTableRowCount
        {
            get { return mMaxTableRowCount; }
            set { mMaxTableRowCount = value; }
        }

        private static string mFontFamily = "Times New Roman";
        public static string FontFamily
        {
            get { return mFontFamily; }
            set { mFontFamily = value; }
        }

        private static float mFontSize = 12.0f;
        public static float FontSize
        {
            get { return mFontSize; }
            set { mFontSize = value; }
        }

        private static float mCaptionFontSize = 9.0f;
        public static float CaptionFontSize
        {
            get { return mCaptionFontSize; }
            set { mCaptionFontSize = value; }
        }

        public static string EncodeName(string content)
        {
            //return content.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;");
            return content;
        }

        public static void AddHeaderLeft(Word.Application myApp, Word.Document myDoc, ref object missing, string header_content)
        {
            foreach (Word.Section section in myDoc.Sections)
            {
                Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Font.ColorIndex = Word.WdColorIndex.wdBlack;
                headerRange.Font.Bold = 1;
                headerRange.Font.Name = mFontFamily;
                headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                headerRange.Text = header_content;
            }
        }

        public static void AddHeaderImageLeft(Word.Application myApp, Word.Document myDoc, ref object missing, string img_filename)
        {
            foreach (Word.Section section in myDoc.Sections)
            {
                Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                headerRange.InlineShapes.AddPicture(img_filename);
            }
        }

        public static void AddFooterPageNumberRight(Word.Application myApp, Word.Document myDoc, ref object missing)
        {
            foreach (Word.Section wordSection in myDoc.Sections)
            {
                Word.Range footerRange = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Font.ColorIndex = Word.WdColorIndex.wdBlack;
                footerRange.Font.Name = mFontFamily;
                footerRange.Font.Size = mFontSize;
                footerRange.Fields.Add(footerRange, Word.WdFieldType.wdFieldPage);
                footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            }
        }

        public static void AddFooterImageLeft(Word.Application myApp, Word.Document myDoc, ref object missing, string img_filename)
        {
            foreach (Word.Section section in myDoc.Sections)
            {
                Word.Range footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                footerRange.InlineShapes.AddPicture(img_filename);
            }
        }

        public static void AddImageFromFile(Word.Application myApp, Word.Document myDoc, ref object missing, string img_filename, string img_caption)
        {
            object oRng = myDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            Word.Paragraph oPara = myDoc.Content.Paragraphs.Add(ref oRng);
            oPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            Word.InlineShape img = oPara.Range.InlineShapes.AddPicture(img_filename);

            oPara.Range.Font.Name = mFontFamily;
            oPara.Range.Font.Size = mCaptionFontSize;
            oPara.Range.Font.Italic = 1;

            oPara.Range.InsertParagraphAfter();

            object oLabel = Word.WdCaptionLabelID.wdCaptionFigure;
            object oTitle = string.Format(" {0}", img_caption);
            object position = Word.WdCaptionPosition.wdCaptionPositionBelow;
            img.Range.InsertCaption(ref oLabel, ref oTitle, ref missing, ref position);
        }

        public static void AddParagraph(Word.Application myApp, Word.Document myDoc, ref object missing, string content)
        {
            object oRng = myDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            Word.Paragraph oPara = myDoc.Content.Paragraphs.Add(ref oRng);
            oPara.Range.Text=content;
            oPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
            oPara.Range.Font.Name = mFontFamily;
            oPara.Range.Font.Size = mFontSize;
            oPara.Format.SpaceAfter = 6;
            oPara.Range.InsertParagraphAfter();  
        }

        public static void AddHeading1(Word.Application myApp, Word.Document myDoc, ref object missing, string content, bool generate_section_number, params int[] section_numbers)
        {
            object oRng = myDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            content = EncodeName(content);
            if (generate_section_number)
            {
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < section_numbers.Length; ++i)
                {
                    if (i != 0)
                    {
                        sb.Append(".");
                    }
                    sb.AppendFormat("{0}", section_numbers[i]);
                }
                sb.AppendFormat(" {0}", content);

                object headingType = Word.WdBuiltinStyle.wdStyleHeading1;

                Word.Paragraph oPara = myDoc.Content.Paragraphs.Add(ref oRng);
                oPara.Range.Text = sb.ToString();
                oPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                oPara.Range.Font.Name = mFontFamily;
                oPara.Range.set_Style(ref headingType);
                oPara.Format.SpaceAfter = 24;
                oPara.Range.InsertParagraphAfter(); 
                
            }
            else
            {
                object headingType = Word.WdBuiltinStyle.wdStyleHeading1;

                Word.Paragraph oPara = myDoc.Content.Paragraphs.Add(ref oRng);
                oPara.Range.Text = content;
                oPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                oPara.Range.Font.Name = mFontFamily;
                oPara.Range.set_Style(ref headingType);
                oPara.Format.SpaceAfter = 24;
                oPara.Range.InsertParagraphAfter(); 
            }
        }

        public static void AddHeading2(Word.Application myApp, Word.Document myDoc, ref object missing, string content, bool generate_section_number, params int[] section_numbers)
        {
            object oRng = myDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            content = EncodeName(content);
            if (generate_section_number)
            {
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < section_numbers.Length; ++i)
                {
                    if (i != 0)
                    {
                        sb.Append(".");
                    }
                    sb.AppendFormat("{0}", section_numbers[i]);
                }
                sb.AppendFormat(" {0}", content);

                object headingType = Word.WdBuiltinStyle.wdStyleHeading2;
                Word.Paragraph oPara = myDoc.Content.Paragraphs.Add(ref oRng);
                oPara.Range.Text = sb.ToString();
                oPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                oPara.Range.Font.Name = mFontFamily;
                oPara.Range.set_Style(ref headingType);
                oPara.Format.SpaceAfter = 24;
                oPara.Range.InsertParagraphAfter(); 
            }
            else
            {
                object headingType = Word.WdBuiltinStyle.wdStyleHeading2;
                Word.Paragraph oPara = myDoc.Content.Paragraphs.Add(ref oRng);
                oPara.Range.Text = content;
                oPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                oPara.Range.Font.Name = mFontFamily;
                oPara.Range.set_Style(ref headingType);
                oPara.Format.SpaceAfter = 24;
                oPara.Range.InsertParagraphAfter(); 
            }
        }

        public static void AddHeading3(Word.Application myApp, Word.Document myDoc, ref object missing, string content, bool generate_section_number, params int[] section_numbers)
        {
            object oRng = myDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            content = EncodeName(content);
            if (generate_section_number)
            {
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < section_numbers.Length; ++i)
                {
                    if (i != 0)
                    {
                        sb.Append(".");
                    }
                    sb.AppendFormat("{0}", section_numbers[i]);
                }
                sb.AppendFormat(" {0}", content);

                object headingType = Word.WdBuiltinStyle.wdStyleHeading3;
                Word.Paragraph oPara = myDoc.Content.Paragraphs.Add(ref oRng);
                oPara.Range.Text = sb.ToString();
                oPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                oPara.Range.Font.Name = mFontFamily;
                oPara.Range.set_Style(ref headingType);
                oPara.Format.SpaceAfter = 24;
                oPara.Range.InsertParagraphAfter(); 
            }
            else
            {
                object headingType = Word.WdBuiltinStyle.wdStyleHeading3;
                Word.Paragraph oPara = myDoc.Content.Paragraphs.Add(ref oRng);
                oPara.Range.Text = content;
                oPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                oPara.Range.Font.Name = mFontFamily;
                oPara.Range.set_Style(ref headingType);
                oPara.Format.SpaceAfter = 24;
                oPara.Range.InsertParagraphAfter(); 
            }
        }

        public static void AddHeading4(Word.Application myApp, Word.Document myDoc, ref object missing, string content, bool generate_section_number, params int[] section_numbers)
        {
            object oRng = myDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            content = EncodeName(content);
            if (generate_section_number)
            {
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < section_numbers.Length; ++i)
                {
                    if (i != 0)
                    {
                        sb.Append(".");
                    }
                    sb.AppendFormat("{0}", section_numbers[i]);
                }
                sb.AppendFormat(" {0}", content);

                object headingType = Word.WdBuiltinStyle.wdStyleHeading4;
                Word.Paragraph oPara = myDoc.Content.Paragraphs.Add(ref oRng);
                oPara.Range.Text = sb.ToString();
                oPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                oPara.Range.Font.Name = mFontFamily;
                oPara.Range.set_Style(ref headingType);
                oPara.Format.SpaceAfter = 24;
                oPara.Range.InsertParagraphAfter(); 
            }
            else
            {
                object headingType = Word.WdBuiltinStyle.wdStyleHeading4;
                Word.Paragraph oPara = myDoc.Content.Paragraphs.Add(ref oRng);
                oPara.Range.Text = content;
                oPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                oPara.Range.Font.Name = mFontFamily;
                oPara.Range.set_Style(ref headingType);
                oPara.Format.SpaceAfter = 24;
                oPara.Range.InsertParagraphAfter(); 
            }
        }

        public static void AddTableFromDataTable(Word.Application myApp, Word.Document myDoc, ref object missing, System.Data.DataTable data_table, string table_caption, EventHandler<ProgressChangedEventArgs> progress_handler)
        {
            int column_count = data_table.Columns.Count;
            int row_count = data_table.Rows.Count;

            DateTime ticked_time = DateTime.Now;
            DateTime interval_time = ticked_time;
            DateTime start_time = ticked_time;

            if (row_count <= mMaxTableRowCount)
            {
                Word.Range oRng = myDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                object oRng2 = myDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

                Word.Paragraph oPara = myDoc.Content.Paragraphs.Add(ref oRng2);
                oPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

                oPara.Range.Font.Name = mFontFamily;
                oPara.Range.Font.Size = mCaptionFontSize;
                oPara.Range.Font.Italic = 1;

                oPara.Range.InsertParagraphAfter();

                Word.Table oTable = oPara.Range.Tables.Add(oRng, row_count + 3, column_count, ref missing, ref missing);

                oTable.set_Style(Word.WdBuiltinStyle.wdStyleTableMediumShading1Accent1);

                oTable.Range.ParagraphFormat.SpaceAfter = 6;

                for (int j = 1; j <= column_count; ++j)
                {
                    string header = EncodeName(data_table.Columns[j - 1].Caption);
                    oTable.Cell(1, j).Range.Text = header;
                }

                for (int i = 0; i < row_count; ++i)
                {
                    DataRow data_row = data_table.Rows[i];
                    for (int j = 1; j <= column_count; ++j)
                    {
                        string column_name = data_table.Columns[j - 1].ColumnName;
                        string cell_value = data_row[column_name].ToString();
                        oTable.Cell(i + 2, j).Range.Text = cell_value;
                    }
                    ticked_time = DateTime.Now;
                    TimeSpan ts = ticked_time - interval_time;
                    if (ts.TotalMilliseconds > 1000)
                    {
                        interval_time = ticked_time;
                        int progress_percentage = i * 100 / row_count;

                        TimeSpan duration = ticked_time - start_time;
                        double duration_in_minutes = duration.TotalMinutes;
                        double remaining_duration_in_minutes = -1;

                        if (progress_percentage > 0)
                        {
                            remaining_duration_in_minutes = duration_in_minutes * (100 - progress_percentage) / progress_percentage;
                        }

                        progress_handler(myApp, new ProgressChangedEventArgs(progress_percentage, string.Format("Create Word Table at Line: {0} of {1} ({2}%) (Dur: {3:0.0} min Remain Dur: {4:0.0} min)", i, row_count, progress_percentage, duration_in_minutes, remaining_duration_in_minutes)));
                    }
                }

                object oLabel = string.Format(" {0}", table_caption);

                object position = Word.WdCaptionPosition.wdCaptionPositionBelow;

                oTable.Range.InsertCaption(Word.WdCaptionLabelID.wdCaptionTable, ref oLabel, ref missing, ref position, ref missing);

                progress_handler(myApp, new ProgressChangedEventArgs(0, "Create Table Completed"));
            }
            else
            {
                int remaining_row_count = row_count;
                int processed_row_count = 0;
                while (remaining_row_count > 0)
                {
                    int actual_row_count = System.Math.Min(mMaxTableRowCount, remaining_row_count);

                    Word.Range oRng = myDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    object oRng2 = myDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

                    Word.Paragraph oPara = myDoc.Content.Paragraphs.Add(ref oRng2);
                    oPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

                    oPara.Range.Font.Name = mFontFamily;
                    oPara.Range.Font.Size = mCaptionFontSize;
                    oPara.Range.Font.Italic = 1;

                    oPara.Range.InsertParagraphAfter();

                    Word.Table oTable = oPara.Range.Tables.Add(oRng, actual_row_count + 3, column_count, ref missing, ref missing);

                    oTable.set_Style(Word.WdBuiltinStyle.wdStyleTableMediumShading1Accent1);

                    oTable.Range.ParagraphFormat.SpaceAfter = 6;

                    for (int j = 1; j <= column_count; ++j)
                    {
                        string header = EncodeName(data_table.Columns[j - 1].Caption);
                        oTable.Cell(1, j).Range.Text = header;
                    }

                    for (int i = 0; i < actual_row_count; ++i)
                    {
                        DataRow data_row = data_table.Rows[processed_row_count + i];
                        for (int j = 1; j <= column_count; ++j)
                        {
                            string column_name = data_table.Columns[j - 1].ColumnName;
                            string cell_value = data_row[column_name].ToString();
                            oTable.Cell(i + 2, j).Range.Text = cell_value;
                        }

                        ticked_time = DateTime.Now;
                        TimeSpan ts = ticked_time - interval_time;
                        if (ts.TotalMilliseconds > 1000)
                        {
                            interval_time = ticked_time;

                            int progress_percentage = (processed_row_count + i) * 100 / row_count;
                            TimeSpan duration = ticked_time - start_time;
                            double duration_in_minutes = duration.TotalMinutes;
                            double remaining_duration_in_minutes = -1;
                            if (progress_percentage > 0)
                            {
                                remaining_duration_in_minutes = duration_in_minutes * (100 - progress_percentage) / progress_percentage;
                            }
                            progress_handler(myApp, new ProgressChangedEventArgs(progress_percentage, string.Format("Create Word Table at Line: {0} of {1} ({2}%) (Dur: {3:0.0} min Remain Dur: {4:0.0} min)", processed_row_count + i, row_count, progress_percentage, duration_in_minutes, remaining_duration_in_minutes)));
                        }
                    }

                    object oLabel = string.Format(" {0}", table_caption);

                    object position = Word.WdCaptionPosition.wdCaptionPositionBelow;

                    oTable.Range.InsertCaption(Word.WdCaptionLabelID.wdCaptionTable, ref oLabel, ref missing, ref position, ref missing);

                    processed_row_count += actual_row_count;
                    remaining_row_count -= actual_row_count;
                }

                progress_handler(myApp, new ProgressChangedEventArgs(0, "Create Table Completed"));
            }

        }
    }
}
