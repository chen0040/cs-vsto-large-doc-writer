using System;
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
            report.StartSection_H1("Section 1");

            ReportGeneratorWithVSTO generator = new ReportGeneratorWithVSTO(report);

            string imageContentFolder = "/tmp";
            generator.GenerateReport("/tmp/hello.doc", imageContentFolder);
        }
    }
}
