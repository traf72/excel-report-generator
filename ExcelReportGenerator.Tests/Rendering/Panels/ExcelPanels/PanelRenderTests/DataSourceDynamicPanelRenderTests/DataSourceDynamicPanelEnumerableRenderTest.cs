using ClosedXML.Excel;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourceDynamicPanelRenderTests
{
    [TestClass]
    public class DataSourceDynamicPanelEnumerableRenderTest
    {
        [TestMethod]
        public void TestRenderEnumerable()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range1 = ws.Range(2, 2, 4, 2);
            range1.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{Headers}";
            ws.Cell(3, 2).Value = "{Data}";
            ws.Cell(4, 2).Value = "{Totals}";

            var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("TestRange"), report, report.TemplateProcessor);
            panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 6, 5), panel.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourceDynamicPanelEnumerableRenderTest),
                nameof(TestRenderEnumerable)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderEmptyEnumerable()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range1 = ws.Range(2, 2, 4, 2);
            range1.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{Headers}";
            ws.Cell(3, 2).Value = "{Data}";
            ws.Cell(4, 2).Value = "{Totals}";

            var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetEmptyIEnumerable()", ws.NamedRange("TestRange"), report, report.TemplateProcessor);
            panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 3, 5), panel.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourceDynamicPanelEnumerableRenderTest),
                nameof(TestRenderEmptyEnumerable)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}