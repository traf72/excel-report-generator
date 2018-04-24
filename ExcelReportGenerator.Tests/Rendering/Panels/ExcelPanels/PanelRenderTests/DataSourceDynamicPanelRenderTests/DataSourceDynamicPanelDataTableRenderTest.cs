using ClosedXML.Excel;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;
using NUnit.Framework;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourceDynamicPanelRenderTests
{
    
    public class DataSourceDynamicPanelDataTableRenderTest
    {
        [Test]
        public void TestRenderDataTable()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 4, 2);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{Headers}";
            ws.Cell(2, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            ws.Cell(2, 2).Style.Border.OutsideBorderColor = XLColor.Red;
            ws.Cell(2, 2).Style.Font.Bold = true;

            ws.Cell(3, 2).Value = "{Data}";
            ws.Cell(3, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Cell(3, 2).Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(4, 2).Value = "{Totals}";
            ws.Cell(4, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Dotted;
            ws.Cell(4, 2).Style.Border.OutsideBorderColor = XLColor.Green;

            var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetAllCustomersDataTable()", ws.NamedRange("TestRange"), report, report.TemplateProcessor);
            panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 6, 7), panel.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourceDynamicPanelDataTableRenderTest),
                nameof(TestRenderDataTable)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [Test]
        public void TestRenderEmptyDataTable()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 4, 2);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{Headers}";
            ws.Cell(2, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            ws.Cell(2, 2).Style.Border.OutsideBorderColor = XLColor.Red;
            ws.Cell(2, 2).Style.Font.Bold = true;

            ws.Cell(3, 2).Value = "{Data}";
            ws.Cell(3, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Cell(3, 2).Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(4, 2).Value = "{Totals}";
            ws.Cell(4, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Dotted;
            ws.Cell(4, 2).Style.Border.OutsideBorderColor = XLColor.Green;

            var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetEmptyDataTable()", ws.NamedRange("TestRange"), report, report.TemplateProcessor);
            panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 3, 7), panel.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourceDynamicPanelDataTableRenderTest),
                nameof(TestRenderEmptyDataTable)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}