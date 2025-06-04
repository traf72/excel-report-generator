using ClosedXML.Excel;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;
using Assert = NUnit.Framework.Legacy.ClassicAssert;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests;

public class TotalsPanelRenderTest
{
    [Test]
    public void TestPanelRender()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");

        var range = ws.Range(1, 1, 1, 12);
        range.AddToNamed("Test", XLScope.Worksheet);

        ws.Cell(1, 1).Value = "Plain text";
        ws.Cell(1, 2).Value = "{Sum(di:Sum)}";
        ws.Cell(1, 3).Value = "{ Custom(DI:Sum, CustomAggregation, PostAggregation)  }";
        ws.Cell(1, 4).Value = "{Min(di:Sum)}";
        ws.Cell(1, 5).Value =
            "Text1 {count(di:Name)} Text2 {avg(di:Sum, , PostAggregationRound)} {p:StrParam} {Max(di:Sum)} {Max(di:Sum)}";
        ws.Cell(1, 6).Value = "{Mix(di:Sum)}";
        ws.Cell(1, 7).FormulaA1 = "=SUM(B1:D1)";
        ws.Cell(1, 8).FormulaA1 = "=ROW()";
        ws.Cell(1, 9).Value = "{sf:Format(p:DateParam, yyyyMMdd)}";
        ws.Cell(1, 10).Value = "{p:IntParam}";
        ws.Cell(1, 11).Value = "{sf:Format(m:TestClassForTotals:Round(Min(di:Sum), 1), #,,0.0000)}";
        ws.Cell(1, 12).Value =
            "Aggregation: {m:TestClassForTotals:Meth(avg(di:Sum, , PostAggregationRound), Max( di : Sum ))}. Date: {sf:Format(p:DateParam, dd.MM.yyyy)}";
        ws.Cell(1, 13).Value = "{Sum(di:Sum)}";

        var panel = new ExcelTotalsPanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("Test"), report,
            report.TemplateProcessor)
        {
            BeforeRenderMethodName = "TestExcelTotalsPanelBeforeRender",
            AfterRenderMethodName = "TestExcelTotalsPanelAfterRender"
        };
        panel.Render();

        Assert.AreEqual(range, panel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(nameof(TotalsPanelRenderTest),
            nameof(TestPanelRender)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void TestPanelRenderWithParentContext()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var parentRange = ws.Range(2, 2, 3, 5);
        parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

        var child1 = ws.Range(2, 2, 2, 5);
        child1.AddToNamed("ChildRange1", XLScope.Worksheet);

        var child2 = ws.Range(3, 2, 3, 5);
        child2.AddToNamed("ChildRange2", XLScope.Worksheet);

        ws.Cell(2, 2).Value = "{di:Field1}";
        ws.Cell(2, 3).Value = "{di:Field2}";

        ws.Cell(3, 2).Value = "{Count(di:Field1)}";
        ws.Cell(3, 3).Value = "{Max(di:Field2)}";
        ws.Cell(3, 4).Value = "{Max(di:parent:Sum)}";
        ws.Cell(3, 5).Value = "{di:parent:Name}";

        var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"),
            report, report.TemplateProcessor);
        var childPanel1 = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.NamedRange("ChildRange1"), report, report.TemplateProcessor)
        {
            Parent = parentPanel
        };
        var childPanel2 = new ExcelTotalsPanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.NamedRange("ChildRange2"), report, report.TemplateProcessor)
        {
            Parent = parentPanel
        };
        parentPanel.Children = new[] {childPanel1, childPanel2};
        parentPanel.Render();

        Assert.AreEqual(ws.Range(2, 2, 9, 5), parentPanel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(nameof(TotalsPanelRenderTest),
            nameof(TestPanelRenderWithParentContext)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void TestPanelWithNoData()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");

        var range = ws.Range(1, 1, 1, 5);
        range.AddToNamed("Test", XLScope.Worksheet);

        ws.Cell(1, 1).Value = "Plain text";
        ws.Cell(1, 2).Value = "{Sum(di:Sum)}";
        ws.Cell(1, 3).Value = "{ Custom(DI:Sum, CustomAggregation, PostAggregation)  }";
        ws.Cell(1, 4).Value = "{Min(di:Sum)}";
        ws.Cell(1, 5).Value = "Text1 {count(di:Name)} Text2 {avg(di:Sum, , PostAggregationRound)} Text3 {Max(di:Sum)}";

        var panel = new ExcelTotalsPanel("m:DataProvider:GetEmptyIEnumerable()", ws.NamedRange("Test"), report,
            report.TemplateProcessor);
        panel.Render();

        Assert.AreEqual(range, panel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(nameof(TotalsPanelRenderTest),
            nameof(TestPanelWithNoData)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void TestPanelWithNullData()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");

        var range = ws.Range(1, 1, 1, 5);
        range.AddToNamed("Test", XLScope.Worksheet);

        ws.Cell(1, 1).Value = "Plain text";
        ws.Cell(1, 2).Value = "{Sum(di:Sum)}";
        ws.Cell(1, 3).Value = "{ Custom(DI:Sum, CustomAggregation, PostAggregation)  }";
        ws.Cell(1, 4).Value = "{Min(di:Sum)}";
        ws.Cell(1, 5).Value = "Text1 {count(di:Name)} Text2 {avg(di:Sum, , PostAggregationRound)} Text3 {Max(di:Sum)}";

        var panel = new ExcelTotalsPanel("m:DataProvider:GetNullItem()", ws.NamedRange("Test"), report,
            report.TemplateProcessor);
        panel.Render();

        Assert.AreEqual(range, panel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(nameof(TotalsPanelRenderTest),
            nameof(TestPanelWithNullData)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    private class TestClassForTotals
    {
        public string Meth(double param1, decimal param2)
        {
            return ((decimal) param1 + param2).ToString("F2");
        }

        public decimal Round(decimal input, int precision)
        {
            return Math.Round(input, precision);
        }
    }
}