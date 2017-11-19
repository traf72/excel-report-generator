using ClosedXML.Excel;
using ExcelReporter.Implementations.Panels.Excel;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;

namespace ExcelReporter.Tests.Implementations.Panels.Excel.PanelRenderTests.DataSourcePanelRenderTests
{
    [TestClass]
    public class DataSourcePanelDataSetRenderTest
    {
        public DataSourcePanelDataSetRenderTest()
        {
            TestHelper.InitDataDirectory();
        }

        [TestMethod]
        public void TestRenderDataSet()
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

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetAllCustomersDataSet()", ws.NamedRange("TestRange"), report);
            panel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelDataSetRenderTest),
                nameof(TestRenderDataSet)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderEmptyDataSet()
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

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetEmptyDataSet()", ws.NamedRange("TestRange"), report);
            panel.Render();

            Assert.AreEqual(0, ws.CellsUsed().Count());

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}