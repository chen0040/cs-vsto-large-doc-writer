# cs-vsto-large-doc-writer
Package provides a large report generator using on word docs using VSTO and C#

# Objective

While there have been quite a number of libraries which claims to be able to generate word document report. However, many of these libraries failed when it came to generate a very large report, one that may contains hundreds of pages or even more. This library was created to enable report generation in word document in these circumstances.

# Langauage

C# 

# Install

The library was built using VS2015 Community Edition. You can clone and build the library then add the library to your references in a .NET project. Note that this library is based on VSTO and thus requires the availability of office 2007 for it to work. It also requires the following COM libraries to be available in the C# project's references

* Microsoft.Office.Core (Version: 2.4)
* Microsoft.Office.Interop.Excel (Version: 1.6)
* Microsoft.Office.Interop.Word (Version: 8.4)

This link below shows how to solve the COM error when uninstall vs 2007 and reinstall some other version of office and then reinstall vs 2007:

https://social.msdn.microsoft.com/Forums/vstudio/en-US/08f13e9d-895c-4102-b6d9-e327af8cf8c0/0x80029c4a-typeecantloadlibrary?forum=vsto

# Usage

Below is the C# sample code for creating a sample report:

```cs 
using LargeDocWriter;

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
```




