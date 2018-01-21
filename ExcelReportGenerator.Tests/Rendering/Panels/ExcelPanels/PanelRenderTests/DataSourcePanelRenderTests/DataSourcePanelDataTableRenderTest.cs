using System.Linq;
using ClosedXML.Excel;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourcePanelRenderTests
{
    [TestClass]
    public class DataSourcePanelDataTableRenderTest
    {
        [TestMethod]
        public void TestRenderDataTable()
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

            var panel = new ExcelDataSourcePanel("m:DataProvider:GetAllCustomersDataTable()", ws.NamedRange("TestRange"), report, report.TemplateProcessor);
            IXLRange resultRange = panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 4, 6), resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelDataTableRenderTest),
                nameof(TestRenderDataTable)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderEmptyDataTable()
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

            var panel = new ExcelDataSourcePanel("m:DataProvider:GetEmptyDataTable()", ws.NamedRange("TestRange"), report, report.TemplateProcessor);
            IXLRange resultRange = panel.Render();

            Assert.IsNull(resultRange);

            Assert.AreEqual(0, ws.CellsUsed().Count());

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}