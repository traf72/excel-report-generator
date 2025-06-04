using ClosedXML.Excel;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;
using Assert = NUnit.Framework.Legacy.ClassicAssert;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourceDynamicPanelRenderTests;

public class DataSourceDynamicPanelEnumerableRenderTest
{
    [Test]
    public void TestRenderEnumerable()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range1 = ws.Range(2, 2, 4, 2);
        range1.AddToNamed("TestRange", XLScope.Worksheet);

        ws.Cell(2, 2).Value = "{Headers}";
        ws.Cell(3, 2).Value = "{Data}";
        ws.Cell(4, 2).Value = "{Totals}";

        var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetIEnumerable()", ws.DefinedName("TestRange"),
            report, report.TemplateProcessor);
        panel.Render();

        Assert.AreEqual(ws.Range(2, 2, 6, 5), panel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(
            nameof(DataSourceDynamicPanelEnumerableRenderTest),
            nameof(TestRenderEnumerable)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void TestRenderEmptyEnumerable()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range1 = ws.Range(2, 2, 4, 2);
        range1.AddToNamed("TestRange", XLScope.Worksheet);

        ws.Cell(2, 2).Value = "{Headers}";
        ws.Cell(3, 2).Value = "{Data}";
        ws.Cell(4, 2).Value = "{Totals}";

        var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetEmptyIEnumerable()", ws.DefinedName("TestRange"),
            report, report.TemplateProcessor);
        panel.Render();

        Assert.AreEqual(ws.Range(2, 2, 3, 5), panel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(
            nameof(DataSourceDynamicPanelEnumerableRenderTest),
            nameof(TestRenderEmptyEnumerable)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }
}