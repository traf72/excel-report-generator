using ClosedXML.Excel;
using ExcelReporter.Enums;
using ExcelReporter.Rendering.Panels.ExcelPanels;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourceDynamicPanelRenderTests
{
    [TestClass]
    public class DataSourceDynamicPanelDataReaderRenderTest
    {
        [TestMethod]
        public void TestRenderDataReader()
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
            ws.Cell(4, 2).Style.NumberFormat.Format = "$ #,##0.00";
            ws.Cell(4, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Dotted;
            ws.Cell(4, 2).Style.Border.OutsideBorderColor = XLColor.Green;

            var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetAllCustomersDataReader()", ws.NamedRange("TestRange"), report);
            panel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourceDynamicPanelDataReaderRenderTest),
                nameof(TestRenderDataReader)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderDataReader_HorizontalPanel()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 2, 4);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{Headers}";
            ws.Cell(2, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            ws.Cell(2, 2).Style.Border.OutsideBorderColor = XLColor.Red;
            ws.Cell(2, 2).Style.Font.Bold = true;

            ws.Cell(2, 3).Value = "{Data}";
            ws.Cell(2, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Cell(2, 3).Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(2, 4).Value = "{Totals}";
            ws.Cell(2, 4).Style.NumberFormat.Format = "$ #,##0.00";
            ws.Cell(2, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Dotted;
            ws.Cell(2, 4).Style.Border.OutsideBorderColor = XLColor.Green;

            var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetAllCustomersDataReader()", ws.NamedRange("TestRange"), report)
            {
                Type = PanelType.Horizontal,
            };
            panel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourceDynamicPanelDataReaderRenderTest),
                nameof(TestRenderDataReader_HorizontalPanel)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderEmptyDataReader()
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
            ws.Cell(4, 2).Style.NumberFormat.Format = "$ #,##0.00";
            ws.Cell(4, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Dotted;
            ws.Cell(4, 2).Style.Border.OutsideBorderColor = XLColor.Green;

            var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetEmptyDataReader()", ws.NamedRange("TestRange"), report);
            panel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourceDynamicPanelDataReaderRenderTest),
                nameof(TestRenderEmptyDataReader)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}