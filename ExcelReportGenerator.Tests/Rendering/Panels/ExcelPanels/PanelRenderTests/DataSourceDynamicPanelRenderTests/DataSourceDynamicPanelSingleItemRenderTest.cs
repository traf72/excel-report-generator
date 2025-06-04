using ClosedXML.Excel;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;
using Assert = NUnit.Framework.Legacy.ClassicAssert;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourceDynamicPanelRenderTests;

public class DataSourceDynamicPanelSingleItemRenderTest
{
    [Test]
    public void TestRenderSingleItem()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range1 = ws.Range(2, 2, 4, 2);
        range1.AddToNamed("TestRange", XLScope.Worksheet);

        ws.Cell(2, 2).Value = "{Headers}";
        ws.Cell(3, 2).Value = "{Data}";
        ws.Cell(4, 2).Value = "{Totals}";

        var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetSingleItem()", ws.DefinedName("TestRange"),
            report, report.TemplateProcessor);
        panel.Render();

        Assert.AreEqual(ws.Range(2, 2, 4, 5), panel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(
            nameof(DataSourceDynamicPanelSingleItemRenderTest),
            nameof(TestRenderSingleItem)), ws.Workbook);

        // report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void TestRenderNullItem()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range1 = ws.Range(2, 2, 4, 2);
        range1.AddToNamed("TestRange", XLScope.Worksheet);

        ws.Cell(2, 2).Value = "{Headers}";
        ws.Cell(3, 2).Value = "{Data}";
        ws.Cell(4, 2).Value = "{Totals}";

        var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetNullItem()", ws.DefinedName("TestRange"), report,
            report.TemplateProcessor);
        panel.Render();

        Assert.IsNull(panel.ResultRange);

        Assert.AreEqual(0, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
    }
}