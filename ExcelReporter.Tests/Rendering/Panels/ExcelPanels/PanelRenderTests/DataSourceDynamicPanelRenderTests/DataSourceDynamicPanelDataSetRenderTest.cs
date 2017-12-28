using ClosedXML.Excel;
using ExcelReporter.Rendering.Panels.ExcelPanels;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourceDynamicPanelRenderTests
{
    [TestClass]
    public class DataSourceDynamicPanelDataSetRenderTest
    {
        [TestMethod]
        public void TestRenderDataSet()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 4, 2);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{Headers}";
            //ws.Cell(2, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            //ws.Cell(2, 2).Style.Border.OutsideBorderColor = XLColor.Red;
            //ws.Cell(2, 2).Style.Font.Bold = true;

            ws.Cell(3, 2).Value = "{Data}";
            //ws.Cell(3, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            //ws.Cell(3, 2).Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(4, 2).Value = "{Totals}";
            //ws.Cell(4, 2).Style.NumberFormat.Format = "$ #,##0.00";
            //ws.Cell(4, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Dotted;
            //ws.Cell(4, 2).Style.Border.OutsideBorderColor = XLColor.Green;

            var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetAllCustomersDataSet()", ws.NamedRange("TestRange"), report)
            {
                BeforeHeadersRenderMethodName = "TestExcelDynamicPanelBeforeHeadersRender",
                AfterHeadersRenderMethodName = "TestExcelDynamicPanelAfterHeadersRender",
                AfterDataTemplatesRenderMethodName = "TestExcelDynamicPanelAfterDataTemplatesRender",
                BeforeDataRenderMethodName = "TestExcelDynamicPanelBeforeDataRender",
                AfterDataItemRenderMethodName = "TestExcelDynamicPanelAfterDataItemRender",
                BeforeTotalsRenderMethodName = "TestExcelDynamicPanelBeforeDataRender",
            };
            panel.Render();

            //ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourceDynamicPanelDataReaderRenderTest),
            //    nameof(TestRenderDataSet)), ws.Workbook);

            report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderEmptyDataSet()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 2);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{Headers}";
            ws.Cell(2, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            ws.Cell(2, 2).Style.Border.OutsideBorderColor = XLColor.Red;
            ws.Cell(2, 2).Style.Font.Bold = true;

            ws.Cell(3, 2).Value = "{Data}";
            ws.Cell(3, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Cell(3, 2).Style.Border.OutsideBorderColor = XLColor.Black;

            var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetEmptyDataSet()", ws.NamedRange("TestRange"), report);
            panel.Render();

            //ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourceDynamicPanelDataReaderRenderTest),
            //    nameof(TestRenderEmptyDataSet)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}