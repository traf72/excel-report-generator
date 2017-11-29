using System.Linq;
using ClosedXML.Excel;
using ExcelReporter.Rendering.Panels.ExcelPanels;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DynamicPanelRenderTests
{
    [TestClass]
    public class DynamicPanelDataReaderRenderTest
    {
        [TestMethod]
        public void TestRenderDataReader()
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

            var panel = new ExcelDynamicPanel("m:DataProvider:GetAllCustomersDataReader()", ws.NamedRange("TestRange"), report);
            panel.Render();

            //ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DynamicPanelDataReaderRenderTest),
            //    nameof(TestRenderDataReader)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}