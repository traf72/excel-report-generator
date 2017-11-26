using System.Linq;
using ClosedXML.Excel;
using ExcelReporter.Rendering.Panels.ExcelPanels;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DynamicPanelRenderTests
{
    [TestClass]
    public class DynamicPanelDataReaderRenderTest
    {
        public DynamicPanelDataReaderRenderTest()
        {
            TestHelper.InitDataDirectory();
        }

        [TestMethod]
        public void TestRenderDataReader()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 2, 2);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            var panel = new ExcelDynamicPanel("m:TestDataProvider:GetAllCustomersDataReader()", ws.NamedRange("TestRange"), report);
            panel.Render();

            //ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelDataReaderRenderTest),
            //    nameof(TestRenderDataReader)), ws.Workbook);

            report.Workbook.SaveAs("test.xlsx");
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