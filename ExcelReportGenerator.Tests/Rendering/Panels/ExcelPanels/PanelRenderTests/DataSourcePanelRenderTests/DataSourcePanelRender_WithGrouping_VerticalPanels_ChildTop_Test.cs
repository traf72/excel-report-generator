using ClosedXML.Excel;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;
using Assert = NUnit.Framework.Legacy.ClassicAssert;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourcePanelRenderTests;

public class DataSourcePanelRender_WithGrouping_VerticalPanels_ChildTop_Test
{
    [Test]
    public void Test_VerticalPanelsGrouping_ChildTop_ParentCellsShiftChildCellsShift()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var parentRange = ws.Range(2, 2, 3, 5);
        parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

        var child = ws.Range(2, 2, 2, 5);
        child.AddToNamed("ChildRange", XLScope.Worksheet);

        child.Range(1, 2, 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        child.Range(1, 2, 1, 4).Style.Border.OutsideBorderColor = XLColor.Red;

        parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

        ws.Cell(3, 2).Value = "{di:Name}";
        ws.Cell(3, 3).Value = "{di:Date}";

        ws.Cell(2, 3).Value = "{di:Field1}";
        ws.Cell(2, 4).Value = "{di:Field2}";
        ws.Cell(2, 5).Value = "{di:parent:Sum}";

        ws.Cell(1, 1).Value = "{di:Name}";
        ws.Cell(4, 1).Value = "{di:Name}";
        ws.Cell(1, 6).Value = "{di:Name}";
        ws.Cell(4, 6).Value = "{di:Name}";
        ws.Cell(3, 1).Value = "{di:Name}";
        ws.Cell(3, 6).Value = "{di:Name}";
        ws.Cell(1, 4).Value = "{di:Name}";
        ws.Cell(4, 4).Value = "{di:Name}";

        var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"),
            report, report.TemplateProcessor);
        var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
        {
            Parent = parentPanel
        };
        parentPanel.Children = new[] {childPanel};
        parentPanel.Render();

        Assert.AreEqual(ws.Range(2, 2, 9, 5), parentPanel.ResultRange);

        ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelRender_WithGrouping_VerticalPanels_ChildTop_Test),
            "ParentCellsShiftChildCellsShift"), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void Test_VerticalPanelsGrouping_ChildTop_ParentRowShiftChildCellsShift()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var parentRange = ws.Range(2, 2, 3, 5);
        parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

        var child = ws.Range(2, 2, 2, 5);
        child.AddToNamed("ChildRange", XLScope.Worksheet);

        child.Range(1, 2, 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        child.Range(1, 2, 1, 4).Style.Border.OutsideBorderColor = XLColor.Red;

        parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

        ws.Cell(3, 2).Value = "{di:Name}";
        ws.Cell(3, 3).Value = "{di:Date}";

        ws.Cell(2, 3).Value = "{di:Field1}";
        ws.Cell(2, 4).Value = "{di:Field2}";
        ws.Cell(2, 5).Value = "{di:parent:Sum}";

        ws.Cell(1, 1).Value = "{di:Name}";
        ws.Cell(4, 1).Value = "{di:Name}";
        ws.Cell(1, 6).Value = "{di:Name}";
        ws.Cell(4, 6).Value = "{di:Name}";
        ws.Cell(3, 1).Value = "{di:Name}";
        ws.Cell(3, 6).Value = "{di:Name}";
        ws.Cell(1, 4).Value = "{di:Name}";
        ws.Cell(4, 4).Value = "{di:Name}";

        var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"),
            report, report.TemplateProcessor)
        {
            ShiftType = ShiftType.Row
        };
        var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
        {
            Parent = parentPanel
        };
        parentPanel.Children = new[] {childPanel};
        parentPanel.Render();

        Assert.AreEqual(ws.Range(2, 2, 9, 5), parentPanel.ResultRange);

        ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelRender_WithGrouping_VerticalPanels_ChildTop_Test),
            "ParentRowShiftChildCellsShift"), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void Test_VerticalPanelsGrouping_ChildTop_ParentRowShiftChildRowShift()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var parentRange = ws.Range(2, 2, 3, 5);
        parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

        var child = ws.Range(2, 2, 2, 5);
        child.AddToNamed("ChildRange", XLScope.Worksheet);

        child.Range(1, 2, 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        child.Range(1, 2, 1, 4).Style.Border.OutsideBorderColor = XLColor.Red;

        parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

        ws.Cell(3, 2).Value = "{di:Name}";
        ws.Cell(3, 3).Value = "{di:Date}";

        ws.Cell(2, 3).Value = "{di:Field1}";
        ws.Cell(2, 4).Value = "{di:Field2}";
        ws.Cell(2, 5).Value = "{di:parent:Sum}";

        ws.Cell(1, 1).Value = "{di:Name}";
        ws.Cell(4, 1).Value = "{di:Name}";
        ws.Cell(1, 6).Value = "{di:Name}";
        ws.Cell(4, 6).Value = "{di:Name}";
        ws.Cell(3, 1).Value = "{di:Name}";
        ws.Cell(3, 6).Value = "{di:Name}";
        ws.Cell(1, 4).Value = "{di:Name}";
        ws.Cell(4, 4).Value = "{di:Name}";

        var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"),
            report, report.TemplateProcessor)
        {
            ShiftType = ShiftType.Row
        };
        var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
        {
            Parent = parentPanel,
            ShiftType = ShiftType.Row
        };
        parentPanel.Children = new[] {childPanel};
        parentPanel.Render();

        Assert.AreEqual(ws.Range(2, 2, 9, 5), parentPanel.ResultRange);

        ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelRender_WithGrouping_VerticalPanels_ChildTop_Test),
            "ParentRowShiftChildRowShift"), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void Test_VerticalPanelsGrouping_ChildTop_ParentNoShiftChildRowShift()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var parentRange = ws.Range(2, 2, 3, 5);
        parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

        var child = ws.Range(2, 2, 2, 5);
        child.AddToNamed("ChildRange", XLScope.Worksheet);

        child.Range(1, 2, 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        child.Range(1, 2, 1, 4).Style.Border.OutsideBorderColor = XLColor.Red;

        parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

        ws.Cell(3, 2).Value = "{di:Name}";
        ws.Cell(3, 3).Value = "{di:Date}";

        ws.Cell(2, 3).Value = "{di:Field1}";
        ws.Cell(2, 4).Value = "{di:Field2}";
        ws.Cell(2, 5).Value = "{di:parent:Sum}";

        ws.Cell(1, 1).Value = "{di:Name}";
        ws.Cell(4, 1).Value = "{di:Name}";
        ws.Cell(1, 6).Value = "{di:Name}";
        ws.Cell(4, 6).Value = "{di:Name}";
        ws.Cell(3, 1).Value = "{di:Name}";
        ws.Cell(3, 6).Value = "{di:Name}";
        ws.Cell(1, 4).Value = "{di:Name}";
        ws.Cell(4, 4).Value = "{di:Name}";

        var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"),
            report, report.TemplateProcessor)
        {
            ShiftType = ShiftType.NoShift
        };
        var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
        {
            Parent = parentPanel,
            ShiftType = ShiftType.Row
        };
        parentPanel.Children = new[] {childPanel};
        parentPanel.Render();

        Assert.AreEqual(ws.Range(2, 2, 9, 5), parentPanel.ResultRange);

        ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelRender_WithGrouping_VerticalPanels_ChildTop_Test),
            "ParentNoShiftChildRowShift"), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void Test_VerticalPanelsGrouping_ChildTop_ParentCellsShiftChildCellsShift_WithFictitiousRow()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var parentRange = ws.Range(2, 2, 4, 5);
        parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

        var child = ws.Range(3, 2, 3, 5);
        child.AddToNamed("ChildRange", XLScope.Worksheet);

        child.Range(1, 2, 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        child.Range(1, 2, 1, 4).Style.Border.OutsideBorderColor = XLColor.Red;

        parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

        ws.Cell(4, 2).Value = "{di:Name}";
        ws.Cell(4, 3).Value = "{di:Date}";

        ws.Cell(3, 3).Value = "{di:Field1}";
        ws.Cell(3, 4).Value = "{di:Field2}";
        ws.Cell(3, 5).Value = "{di:parent:Sum}";

        ws.Cell(1, 1).Value = "{di:Name}";
        ws.Cell(5, 1).Value = "{di:Name}";
        ws.Cell(1, 6).Value = "{di:Name}";
        ws.Cell(5, 6).Value = "{di:Name}";
        ws.Cell(3, 1).Value = "{di:Name}";
        ws.Cell(3, 6).Value = "{di:Name}";
        ws.Cell(1, 4).Value = "{di:Name}";
        ws.Cell(5, 4).Value = "{di:Name}";

        var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"),
            report, report.TemplateProcessor);
        var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
        {
            Parent = parentPanel
        };
        parentPanel.Children = new[] {childPanel};
        parentPanel.Render();

        Assert.AreEqual(ws.Range(2, 2, 12, 5), parentPanel.ResultRange);

        ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelRender_WithGrouping_VerticalPanels_ChildTop_Test),
            "ParentCellsShiftChildCellsShift_WithFictitiousRow"), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void Test_VerticalPanelsGrouping_ChildTop_ParentCellsShiftChildRowShift_WithFictitiousRow()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var parentRange = ws.Range(2, 2, 4, 5);
        parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

        var child = ws.Range(3, 2, 3, 5);
        child.AddToNamed("ChildRange", XLScope.Worksheet);

        child.Range(1, 2, 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        child.Range(1, 2, 1, 4).Style.Border.OutsideBorderColor = XLColor.Red;

        parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

        ws.Cell(4, 2).Value = "{di:Name}";
        ws.Cell(4, 3).Value = "{di:Date}";

        ws.Cell(3, 3).Value = "{di:Field1}";
        ws.Cell(3, 4).Value = "{di:Field2}";
        ws.Cell(3, 5).Value = "{di:parent:Sum}";

        ws.Cell(1, 1).Value = "{di:Name}";
        ws.Cell(5, 1).Value = "{di:Name}";
        ws.Cell(1, 6).Value = "{di:Name}";
        ws.Cell(5, 6).Value = "{di:Name}";
        ws.Cell(3, 1).Value = "{di:Name}";
        ws.Cell(3, 6).Value = "{di:Name}";
        ws.Cell(1, 4).Value = "{di:Name}";
        ws.Cell(5, 4).Value = "{di:Name}";

        var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"),
            report, report.TemplateProcessor);
        var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
        {
            Parent = parentPanel,
            ShiftType = ShiftType.Row
        };
        parentPanel.Children = new[] {childPanel};
        parentPanel.Render();

        Assert.AreEqual(ws.Range(2, 2, 12, 5), parentPanel.ResultRange);

        ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelRender_WithGrouping_VerticalPanels_ChildTop_Test),
            "ParentCellsShiftChildRowShift_WithFictitiousRow"), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void
        Test_VerticalPanelsGrouping_ChildTop_ParentNoShiftChildCellsShift_WithFictitiousRowWhichDeleteAfterRender()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var parentRange = ws.Range(2, 2, 4, 5);
        parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

        var child = ws.Range(3, 2, 3, 5);
        child.AddToNamed("ChildRange", XLScope.Worksheet);

        child.Range(1, 2, 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        child.Range(1, 2, 1, 4).Style.Border.OutsideBorderColor = XLColor.Red;

        parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

        ws.Cell(4, 2).Value = "{di:Name}";
        ws.Cell(4, 3).Value = "{di:Date}";

        ws.Cell(3, 3).Value = "{di:Field1}";
        ws.Cell(3, 4).Value = "{di:Field2}";
        ws.Cell(3, 5).Value = "{di:parent:Sum}";

        ws.Cell(1, 1).Value = "{di:Name}";
        ws.Cell(5, 1).Value = "{di:Name}";
        ws.Cell(1, 6).Value = "{di:Name}";
        ws.Cell(5, 6).Value = "{di:Name}";
        ws.Cell(3, 1).Value = "{di:Name}";
        ws.Cell(3, 6).Value = "{di:Name}";
        ws.Cell(1, 4).Value = "{di:Name}";
        ws.Cell(5, 4).Value = "{di:Name}";

        var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"),
            report, report.TemplateProcessor)
        {
            AfterDataItemRenderMethodName = "AfterRenderParentDataSourcePanelChildTop",
            ShiftType = ShiftType.NoShift
        };
        var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
        {
            Parent = parentPanel
        };
        parentPanel.Children = new[] {childPanel};
        parentPanel.Render();

        Assert.AreEqual(ws.Range(2, 2, 9, 5), parentPanel.ResultRange);

        ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelRender_WithGrouping_VerticalPanels_ChildTop_Test),
            "ParentNoShiftChildCellsShift_WithFictitiousRowWhichDeleteAfterRender"), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }
}