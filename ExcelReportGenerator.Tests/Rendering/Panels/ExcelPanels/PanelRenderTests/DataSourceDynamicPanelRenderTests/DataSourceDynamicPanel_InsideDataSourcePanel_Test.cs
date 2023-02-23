using ClosedXML.Excel;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourceDynamicPanelRenderTests;

public class DataSourceDynamicPanel_InsideDataSourcePanel_Test
{
    [Test]
    public void TestRender_DynamicPanel_In_DataSourcePanel_Vertical()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var parentRange = ws.Range(1, 2, 4, 2);
        parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

        var childRange = ws.Range(2, 2, 4, 2);
        childRange.AddToNamed("ChildRange", XLScope.Worksheet);

        ws.Cell(1, 2).Value = "{di:Name}";

        ws.Cell(2, 2).Value = "{Headers}";
        ws.Cell(3, 2).Value = "{Data}";
        ws.Cell(4, 2).Value = "{Totals}";

        var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"),
            report, report.TemplateProcessor);
        var childPanel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
        {
            Parent = parentPanel
        };
        parentPanel.Children = new[] {childPanel};
        parentPanel.Render();

        Assert.AreEqual(ws.Range(1, 2, 12, 3), parentPanel.ResultRange);

        ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(
            nameof(DataSourceDynamicPanel_InsideDataSourcePanel_Test),
            nameof(TestRender_DynamicPanel_In_DataSourcePanel_Vertical)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void TestRender_DynamicPanel_In_DataSourcePanel_Horizontal()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var parentRange = ws.Range(2, 1, 2, 4);
        parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

        var childRange = ws.Range(2, 2, 2, 4);
        childRange.AddToNamed("ChildRange", XLScope.Worksheet);

        ws.Cell(2, 1).Value = "{di:Name}";

        ws.Cell(2, 2).Value = "{Headers}";
        ws.Cell(2, 3).Value = "{Numbers}";
        ws.Cell(2, 4).Value = "{Data}";

        var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"),
            report, report.TemplateProcessor)
        {
            Type = PanelType.Horizontal
        };
        var childPanel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
        {
            Parent = parentPanel,
            Type = PanelType.Horizontal
        };
        parentPanel.Children = new[] {childPanel};
        parentPanel.Render();

        Assert.AreEqual(ws.Range(2, 1, 3, 12), parentPanel.ResultRange);

        ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(
            nameof(DataSourceDynamicPanel_InsideDataSourcePanel_Test),
            nameof(TestRender_DynamicPanel_In_DataSourcePanel_Horizontal)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }
}