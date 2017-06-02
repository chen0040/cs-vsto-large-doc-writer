using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using LargeDocWriter.Elements;

namespace LargeDocWriter
{
    public class ReportModel
    {
        protected List<ReportSection> mSections = new List<ReportSection>();
        protected ReportSection mCurrentSection = null;
        protected int mFigureCounter = 0;
        protected string mLeftHeaderContent = "Ministry of Labor - Kingdom of Saudi Arabia";
        protected string mLeftHeaderImagePath = null;

        public string LeftHeaderContent
        {
            get { return mLeftHeaderContent; }
            set { mLeftHeaderContent = value; }
        }

        public string LeftHeaderImagePath
        {
            get { return mLeftHeaderImagePath; }
            set { mLeftHeaderImagePath = value; }
        }

        public List<ReportSection> Sections
        {
            get { return mSections; }
        }

        public ReportModel()
        {

        }

        public void Reset()
        {
            mSections.Clear();
            mFigureCounter = 0;
            mCurrentSection = null;
        }

        public void AppendFigure(string filename, string caption, string tag, int width, int height)
        {
            ReportFigure figure = new ReportFigure(mCurrentSection);
            figure.FileName = filename;
            figure.Caption = caption;
            figure.Tag = tag;
            figure.Width = width;
            figure.Height = height;
            mCurrentSection.AddFigure(figure);
        }

        public void AppendPieChart(Dictionary<string, float> table, string figure_caption, string figure_title, string figure_tag, int chart_width, int chart_height, string figure_filename)
        {
            ReportFigure_PieChart chart = new ReportFigure_PieChart(mCurrentSection);
            chart.Caption = figure_caption;
            chart.Title = figure_title;
            chart.Tag = figure_tag;
            chart.FileName = figure_filename;

            chart.Width = chart_width;
            chart.Height = chart_height;

            chart.Content = table;

            mCurrentSection.AddFigure(chart);
        }

        public void AppendColumnChart(Dictionary<string, float> table, string x_label, string y_label, string figure_caption, string figure_title, string figure_tag, int chart_width, int chart_height, string figure_filename)
        {
            ReportFigure_ColumnChart chart = new ReportFigure_ColumnChart(mCurrentSection);
            chart.Caption = figure_caption;
            chart.Title = figure_title;
            chart.Tag = figure_tag;
            chart.FileName = figure_filename;

            chart.Width = chart_width;
            chart.Height = chart_height;

            chart.XLabel = x_label;
            chart.YLabel = y_label;

            chart.Content = table;

            mCurrentSection.AddFigure(chart);
        }

        public void AppendBarChart(Dictionary<string, float> table, string figure_caption, string figure_title, string figure_tag, int chart_width, int chart_height, string figure_filename)
        {
            ReportFigure_BarChart chart = new ReportFigure_BarChart(mCurrentSection);
            chart.Caption = figure_caption;
            chart.Title = figure_title;
            chart.Tag = figure_tag;
            chart.FileName = figure_filename;

            chart.Width = chart_width;
            chart.Height = chart_height;

            chart.Content = table;

            mCurrentSection.AddFigure(chart);
        }

        public void AppendTable(Dictionary<string, float> table_content, string table_caption, string[] headers, string tag)
        {
            ReportTable table = new ReportTable(mCurrentSection);
            DataTable data = new DataTable();

            DataColumn c=new DataColumn();
            c.DataType = typeof(string);
            c.ColumnName = headers[0];
            c.Caption = headers[0];
            data.Columns.Add(c);

            c = new DataColumn();
            c.DataType = typeof(float);
            c.ColumnName = headers[1];
            c.Caption = headers[1];
            data.Columns.Add(c);

            foreach (KeyValuePair<string, float> entry in table_content)
            {
                DataRow row=data.NewRow();
                row[0] = entry.Key;
                row[1] = entry.Value;
                data.Rows.Add(row);
            }

            table.Content = data;
            table.Caption = table_caption;
            table.Headers = headers;
            table.Tag = tag;
            mCurrentSection.AddTable(table);
        }

        public void AppendTable(DataTable db_table, string table_caption, string tag)
        {
            ReportTable table = new ReportTable(mCurrentSection);

            int column_count=db_table.Columns.Count;
            string[] headers = new string[column_count];

            for (int i = 0; i < column_count; ++i )
            {
                headers[i] = db_table.Columns[i].Caption;
            }

            table.Caption = table_caption;
            table.Headers = headers;
            table.Tag = tag;
            table.Content = db_table;
            mCurrentSection.AddTable(table);
        }

        public void AppendTable(Dictionary<string, int> table_content, string table_caption, string[] headers, string tag)
        {
            ReportTable table = new ReportTable(mCurrentSection);

            DataTable data = new DataTable();

            DataColumn c = new DataColumn();
            c.DataType = typeof(string);
            c.ColumnName = headers[0];
            c.Caption = headers[0];
            data.Columns.Add(c);

            c = new DataColumn();
            c.DataType = typeof(int);
            c.ColumnName = headers[1];
            c.Caption = headers[1];
            data.Columns.Add(c);

            foreach (KeyValuePair<string, int> entry in table_content)
            {
                DataRow row = data.NewRow();
                row[0] = entry.Key;
                row[1] = entry.Value;
                data.Rows.Add(row);
            }

            table.Caption = table_caption;
            table.Headers = headers;
            table.Tag = tag;
            table.Content = data;
            mCurrentSection.AddTable(table);
        }

        public void AppendTable(Dictionary<string, List<float>> table_content, string table_caption, string[] headers, string tag)
        {
            ReportTable table = new ReportTable(mCurrentSection);

            DataTable data = new DataTable();

            DataColumn c = new DataColumn(headers[0], typeof(string));
            data.Columns.Add(c);

            int last_cindex = table_content.First().Value.Count;
            for (int cindex = 1; cindex <= last_cindex; ++cindex )
            {
                c = new DataColumn(headers[cindex], typeof(float));
                data.Columns.Add(c);
            }

            foreach(KeyValuePair<string, List<float>> entry in table_content)
            {
                string key = entry.Key;
                List<float> value = entry.Value;
                DataRow row = data.NewRow();
                
                row[0] = key;
                for (int cindex = 1; cindex <= last_cindex; ++cindex)
                {
                    row[cindex] = value[cindex - 1];
                }
                data.Rows.Add(row);
            }

            table.Content = data;
            table.Caption = table_caption;
            table.Headers = headers;
            table.Tag = tag;
            mCurrentSection.AddTable(table);
        }

        public string GenerateFigureFileName()
        {
            string figure_filename = string.Format("{0:0000}.png", mFigureCounter);
            mFigureCounter++;
            return figure_filename;
        }

        public void AppendParagraph(string content, params object[] args)
        {
            ReportParagraph p = new ReportParagraph(mCurrentSection, content, args);
            mCurrentSection.AddParagraph(p);
        }

        public void StartSection_H1(string content)
        {
            int level = 1;
            ReportSection section_h1 = new ReportSection(content, level);
            mSections.Add(section_h1);
            mCurrentSection = section_h1;
        }

        public void StartSection_H2(string content)
        {
            int level = 2;
            StartSubSection(content, level);   
        }

        public void StartSection_H3(string content)
        {
            int level = 3;
            StartSubSection(content, level);
        }

        public void StartSection_H4(string content)
        {
            int level = 4;
            StartSubSection(content, level);
        }

        public void StartSubSection(string content, int level)
        {
            ReportSection section_h2 = new ReportSection(content, level);
            while (mCurrentSection.Level >= level)
            {
                mCurrentSection = mCurrentSection.ParentSection;
            }
            mCurrentSection.AddSubSection(section_h2);
            mCurrentSection = section_h2;
        }

        public string GenerateAnchor(string name, string tag)
        {
            return string.Format("[Anchor|Tag:{0}|Name:{1}]", tag, name);
        }
    }
}
