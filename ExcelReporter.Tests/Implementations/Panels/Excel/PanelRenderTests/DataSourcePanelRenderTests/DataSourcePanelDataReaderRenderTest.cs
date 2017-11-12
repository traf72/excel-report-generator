using System.Linq;
using ClosedXML.Excel;
using ExcelReporter.Implementations.Panels.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Implementations.Panels.Excel.PanelRenderTests.DataSourcePanelRenderTests
{
    [TestClass]
    public class DataSourcePanelDataReaderRenderTest
    {
        public DataSourcePanelDataReaderRenderTest()
        {
            TestHelper.InitDataDirectory();
        }

        [TestMethod]
        public void TestRenderDataReader()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 2, 6);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Id}";
            ws.Cell(2, 3).Value = "{di:Name}";
            ws.Cell(2, 4).Value = "{di:IsVip}";
            ws.Cell(2, 5).Value = "{di:Description}";
            ws.Cell(2, 6).Value = "{di:Type}";

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetAllCustomersDataReader()", ws.NamedRange("TestRange"), report);
            panel.Render();

            Assert.AreEqual(11, ws.CellsUsed().Count());
            Assert.AreEqual(1d, ws.Cell(2, 2).Value);
            Assert.AreEqual("Customer 1", ws.Cell(2, 3).Value);
            Assert.AreEqual(false, ws.Cell(2, 4).Value);
            Assert.AreEqual(string.Empty, ws.Cell(2, 5).Value);
            Assert.AreEqual(string.Empty, ws.Cell(2, 6).Value);
            Assert.AreEqual(2d, ws.Cell(3, 2).Value);
            Assert.AreEqual("Customer 2", ws.Cell(3, 3).Value);
            Assert.AreEqual(true, ws.Cell(3, 4).Value);
            Assert.AreEqual("Reliable", ws.Cell(3, 5).Value);
            Assert.AreEqual(1d, ws.Cell(3, 6).Value);
            Assert.AreEqual(3d, ws.Cell(4, 2).Value);
            Assert.AreEqual("Customer 3", ws.Cell(4, 3).Value);
            Assert.AreEqual(string.Empty, ws.Cell(4, 4).Value);
            Assert.AreEqual("Lost", ws.Cell(4, 5).Value);
            Assert.AreEqual(string.Empty, ws.Cell(4, 6).Value);

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderEmptyDataReader()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 2, 6);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Id}";
            ws.Cell(2, 3).Value = "{di:Name}";
            ws.Cell(2, 4).Value = "{di:IsVip}";
            ws.Cell(2, 5).Value = "{di:Description}";
            ws.Cell(2, 6).Value = "{di:Type}";

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetEmptyDataReader()", ws.NamedRange("TestRange"), report);
            panel.Render();

            Assert.AreEqual(0, ws.CellsUsed().Count());

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}