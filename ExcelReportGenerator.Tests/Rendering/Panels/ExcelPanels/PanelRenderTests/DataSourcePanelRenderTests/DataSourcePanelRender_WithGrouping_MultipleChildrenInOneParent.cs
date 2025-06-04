using ClosedXML.Excel;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;
using Assert = NUnit.Framework.Legacy.ClassicAssert;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourcePanelRenderTests;

public class DataSourcePanelRender_WithGrouping_MultipleChildrenInOneParent
{
    [Test]
    public void Test_TwoChildren_Vertical()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var parentRange = ws.Range(2, 2, 6, 5);
        parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

        var child1 = ws.Range(4, 2, 4, 5);
        child1.AddToNamed("ChildRange1", XLScope.Worksheet);

        var child2 = ws.Range(6, 2, 6, 5);
        child2.AddToNamed("ChildRange2", XLScope.Worksheet);

        ws.Cell(2, 2).Value = "{di:Name}";

        ws.Cell(3, 3).Value = "Field1";
        ws.Cell(3, 3).Style.Font.Bold = true;
        ws.Cell(3, 4).Value = "Field2";
        ws.Cell(3, 4).Style.Font.Bold = true;
        ws.Cell(4, 3).Value = "{di:Field1}";
        ws.Cell(4, 4).Value = "{di:Field2}";
        ws.Cell(5, 5).Value = "Number";
        ws.Cell(5, 5).Style.Font.Bold = true;
        ws.Cell(6, 5).Value = "{di:di}";

        var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.DefinedName("ParentRange"),
            report, report.TemplateProcessor);
        var childPanel1 = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.DefinedName("ChildRange1"), report, report.TemplateProcessor)
        {
            Parent = parentPanel
        };
        var childPanel2 = new ExcelDataSourcePanel("di:ChildrenPrimitive", ws.DefinedName("ChildRange2"), report,
            report.TemplateProcessor)
        {
            Parent = parentPanel
        };
        parentPanel.Children = new[] {childPanel1, childPanel2};
        parentPanel.Render();

        Assert.AreEqual(ws.Range(2, 2, 21, 5), parentPanel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelRender_WithGrouping_MultipleChildrenInOneParent),
            nameof(Test_TwoChildren_Vertical)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void Test_TwoChildren_Horizontal()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var parentRange = ws.Range(2, 2, 5, 6);
        parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

        var child1 = ws.Range(2, 4, 5, 4);
        child1.AddToNamed("ChildRange1", XLScope.Worksheet);

        var child2 = ws.Range(2, 6, 5, 6);
        child2.AddToNamed("ChildRange2", XLScope.Worksheet);

        ws.Cell(2, 2).Value = "{di:Name}";

        ws.Cell(3, 3).Value = "Field1";
        ws.Cell(3, 3).Style.Font.Bold = true;
        ws.Cell(4, 3).Value = "Field2";
        ws.Cell(4, 3).Style.Font.Bold = true;
        ws.Cell(3, 4).Value = "{di:Field1}";
        ws.Cell(4, 4).Value = "{di:Field2}";
        ws.Cell(5, 5).Value = "Number";
        ws.Cell(5, 5).Style.Font.Bold = true;
        ws.Cell(5, 6).Value = "{di:di}";

        var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.DefinedName("ParentRange"),
            report, report.TemplateProcessor);
        var childPanel1 = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.DefinedName("ChildRange1"), report, report.TemplateProcessor)
        {
            Parent = parentPanel,
            Type = PanelType.Horizontal
        };
        var childPanel2 = new ExcelDataSourcePanel("di:ChildrenPrimitive", ws.DefinedName("ChildRange2"), report,
            report.TemplateProcessor)
        {
            Parent = parentPanel,
            Type = PanelType.Horizontal
        };
        parentPanel.Children = new[] {childPanel1, childPanel2};
        parentPanel.Render();

        Assert.AreEqual(ws.Range(2, 2, 13, 8), parentPanel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelRender_WithGrouping_MultipleChildrenInOneParent),
            nameof(Test_TwoChildren_Horizontal)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }
}