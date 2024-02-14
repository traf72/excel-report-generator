using ClosedXML.Excel;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;
using Assert = NUnit.Framework.Legacy.ClassicAssert;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourcePanelRenderTests;

public class DataSourcePanelRender_WithGrouping_ChildEqualsParent_Test
{
    [Test]
    public void Test_VerticalPanelsGrouping_ChildEqualsParent()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var parentRange = ws.Range(2, 2, 2, 4);
        parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

        var child = ws.Range(2, 2, 2, 4);
        child.AddToNamed("ChildRange", XLScope.Worksheet);

        ws.Cell(2, 2).Value = "{di:Field1}";
        ws.Cell(2, 3).Value = "{di:Field2}";
        ws.Cell(2, 4).Value = "{di:parent:Sum}";

        var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"),
            report, report.TemplateProcessor);
        var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
        {
            Parent = parentPanel
        };
        parentPanel.Children = new[] {childPanel};
        parentPanel.Render();

        Assert.AreEqual(ws.Range(2, 2, 6, 4), parentPanel.ResultRange);

        ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelRender_WithGrouping_ChildEqualsParent_Test),
            "Test_VerticalPanelsGrouping_ChildEqualsParent"), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }
}