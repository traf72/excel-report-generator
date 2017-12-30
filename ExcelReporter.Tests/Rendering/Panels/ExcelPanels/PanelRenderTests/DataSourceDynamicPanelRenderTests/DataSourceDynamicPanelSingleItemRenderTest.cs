using System.Linq;
using ClosedXML.Excel;
using ExcelReporter.Rendering.Panels.ExcelPanels;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourceDynamicPanelRenderTests
{
    [TestClass]
    public class DataSourceDynamicPanelSingleItemRenderTest
    {
        [TestMethod]
        public void TestRenderSingleItem()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range1 = ws.Range(2, 2, 4, 2);
            range1.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{Headers}";
            ws.Cell(3, 2).Value = "{Data}";
            ws.Cell(4, 2).Value = "{Totals}";

            var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetSingleItem()", ws.NamedRange("TestRange"), report);
            panel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourceDynamicPanelSingleItemRenderTest),
                nameof(TestRenderSingleItem)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderNullItem()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range1 = ws.Range(2, 2, 4, 2);
            range1.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{Headers}";
            ws.Cell(3, 2).Value = "{Data}";
            ws.Cell(4, 2).Value = "{Totals}";

            var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetNullItem()", ws.NamedRange("TestRange"), report);
            panel.Render();

            Assert.AreEqual(0, ws.CellsUsed().Count());
        }
    }
}