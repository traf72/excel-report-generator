using ClosedXML.Excel;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests
{
    [TestClass]
    public class TotalsPanelRenderTest
    {
        [TestMethod]
        public void TestPanelRender()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");

            IXLRange range = ws.Range(1, 1, 1, 8);
            range.AddToNamed("Test", XLScope.Worksheet);

            ws.Cell(1, 1).Value = "Plain text";
            ws.Cell(1, 2).Value = "{Sum(di:Sum)}";
            ws.Cell(1, 3).Value = "{ Custom(DI:Sum, CustomAggregation, PostAggregation)  }";
            ws.Cell(1, 4).Value = "{Min(di:Sum)}";
            ws.Cell(1, 5).Value = "Text1 {count(Name)} Text2 {avg(di:Sum, , PostAggregationRound)} Text3 {Max(Sum)}";
            ws.Cell(1, 6).Value = "{Mix(di:Sum)}";
            ws.Cell(1, 7).FormulaA1 = "=SUM(B1:D1)";
            ws.Cell(1, 8).FormulaA1 = "=ROW()";
            ws.Cell(1, 9).Value = "{Sum(di:Sum)}";

            var panel = new ExcelTotalsPanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("Test"), report, report.TemplateProcessor)
            {
                BeforeRenderMethodName = "TestExcelTotalsPanelBeforeRender",
                AfterRenderMethodName = "TestExcelTotalsPanelAfterRender",
            };
            IXLRange resultRange = panel.Render();

            Assert.AreEqual(range, resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(TotalsPanelRenderTest),
                nameof(TestPanelRender)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestPanelWithNoData()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");

            IXLRange range = ws.Range(1, 1, 1, 5);
            range.AddToNamed("Test", XLScope.Worksheet);

            ws.Cell(1, 1).Value = "Plain text";
            ws.Cell(1, 2).Value = "{Sum(di:Sum)}";
            ws.Cell(1, 3).Value = "{ Custom(DI:Sum, CustomAggregation, PostAggregation)  }";
            ws.Cell(1, 4).Value = "{Min(di:Sum)}";
            ws.Cell(1, 5).Value = "Text1 {count(Name)} Text2 {avg(di:Sum, , PostAggregationRound)} Text3 {Max(Sum)}";

            var panel = new ExcelTotalsPanel("m:DataProvider:GetEmptyIEnumerable()", ws.NamedRange("Test"), report, report.TemplateProcessor);
            IXLRange resultRange = panel.Render();

            Assert.AreEqual(range, resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(TotalsPanelRenderTest),
                nameof(TestPanelWithNoData)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestPanelWithNullData()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");

            IXLRange range = ws.Range(1, 1, 1, 5);
            range.AddToNamed("Test", XLScope.Worksheet);

            ws.Cell(1, 1).Value = "Plain text";
            ws.Cell(1, 2).Value = "{Sum(di:Sum)}";
            ws.Cell(1, 3).Value = "{ Custom(DI:Sum, CustomAggregation, PostAggregation)  }";
            ws.Cell(1, 4).Value = "{Min(di:Sum)}";
            ws.Cell(1, 5).Value = "Text1 {count(Name)} Text2 {avg(di:Sum, , PostAggregationRound)} Text3 {Max(Sum)}";

            var panel = new ExcelTotalsPanel("m:DataProvider:GetNullItem()", ws.NamedRange("Test"), report, report.TemplateProcessor);
            IXLRange resultRange = panel.Render();

            Assert.AreEqual(range, resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(TotalsPanelRenderTest),
                nameof(TestPanelWithNullData)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}