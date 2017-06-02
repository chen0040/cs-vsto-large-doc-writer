using System;
using System.Collections.Generic;
using System.Data;
using LargeDocWriter;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace cs_vsto_large_doc_writer_unit_test
{
    [TestClass]
    public class ReportGeneratorWithVSTOUnitTest
    {
        [TestMethod]
        public void TestReportWriterMethod()
        {
            ReportModel report = new ReportModel();
            report.StartSection_H1("Section One");

            report.AppendParagraph("Hello this is some text");

            report.StartSection_H2("Paragraph Illustration");

            report.AppendParagraph("Hello this is second paragraph");

            report.AppendParagraph("Hello this is third paragraph");

            //report.AppendFigure("some_picture.jpg", "Some figure", 500, 300);

            report.StartSection_H2("Table Illustration");

            report.AppendParagraph("This section shows how to create table.");

            DataTable table = new DataTable();
            table.Columns.Add("Column1");
            table.Columns.Add("Column2");

            DataRow row = table.NewRow();
            row["Column1"] = 200;
            row["Column2"] = 500;
            table.Rows.Add(row);

            report.AppendTable(table, "Some table");

            report.StartSection_H2("Chart Illustration");

            report.AppendParagraph("This section shows how to create charts.");

            Dictionary<string, float> barData = new Dictionary<string, float>();
            barData["one"] = 200;
            barData["two"] = 500;
            barData["three"] = 300;
            
            report.AppendBarChart(barData, "Some sample data", 500, 300);

            report.AppendColumnChart(barData, "Some column bar", 500, 300);

            report.AppendPieChart(barData, "some pie chart", 400, 400);

            report.StartSection_H1("Conclusion");

            report.AppendParagraph("Some conclusion");
            

            ReportGeneratorWithVSTO generator = new ReportGeneratorWithVSTO(report);
            
            string imageContentFolder = "/tmp";
            generator.GenerateReport("/tmp/hello.doc", imageContentFolder);
        }
    }
}
