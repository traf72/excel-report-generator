﻿using ClosedXML.Excel;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;
using Assert = NUnit.Framework.Legacy.ClassicAssert;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourcePanelRenderTests;

public class DataSourcePanelRender_WithGrouping_HorizontalPanels_ChildLeft_Test
{
    [Test]
    public void Test_HorizontalPanelsGrouping_ChildLeft_ParentCellsShiftChildCellsShift()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var parentRange = ws.Range(2, 2, 5, 3);
        parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

        var child = ws.Range(2, 2, 5, 2);
        child.AddToNamed("ChildRange", XLScope.Worksheet);

        child.Range(2, 1, 4, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        child.Range(2, 1, 4, 1).Style.Border.OutsideBorderColor = XLColor.Red;

        parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

        ws.Cell(2, 3).Value = "{di:Name}";
        ws.Cell(3, 3).Value = "{di:Date}";

        ws.Cell(3, 2).Value = "{di:Field1}";
        ws.Cell(4, 2).Value = "{di:Field2}";
        ws.Cell(5, 2).Value = "{di:parent:Sum}";

        ws.Cell(1, 1).Value = "{di:Name}";
        ws.Cell(1, 3).Value = "{di:Name}";
        ws.Cell(1, 4).Value = "{di:Name}";
        ws.Cell(3, 1).Value = "{di:Name}";
        ws.Cell(3, 4).Value = "{di:Name}";
        ws.Cell(6, 1).Value = "{di:Name}";
        ws.Cell(6, 3).Value = "{di:Name}";
        ws.Cell(6, 4).Value = "{di:Name}";

        var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"),
            report, report.TemplateProcessor)
        {
            Type = PanelType.Horizontal
        };
        var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
        {
            Parent = parentPanel,
            Type = PanelType.Horizontal
        };
        parentPanel.Children = new[] {childPanel};
        parentPanel.Render();

        Assert.AreEqual(ws.Range(2, 2, 5, 9), parentPanel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelRender_WithGrouping_HorizontalPanels_ChildLeft_Test),
            "ParentCellsShiftChildCellsShift"), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void Test_HorizontalPanelsGrouping_ChildLeft_ParentCellsShiftChildCellsShift_WithRichText()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var parentRange = ws.Range(2, 2, 5, 3);
        parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

        var child = ws.Range(2, 2, 5, 2);
        child.AddToNamed("ChildRange", XLScope.Worksheet);

        child.Range(2, 1, 4, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        child.Range(2, 1, 4, 1).Style.Border.OutsideBorderColor = XLColor.Red;

        parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

        ws.Cell(2, 3).CreateRichText()
            .AddText("Name - {di:Name} - {di:Name} - Name")
            .SetItalic();

        ws.Cell(2, 3)
            .GetRichText()
            .Substring(0, 4)
            .SetFontColor(XLColor.Green);
        
        ws.Cell(2, 3)
            .GetRichText()
            .Substring(6, 23)
            .SetFontColor(XLColor.Red);
        
        ws.Cell(2, 3)
            .GetRichText()
            .Substring(31, 4)
            .SetFontColor(XLColor.Blue);

        ws.Cell(3, 3).CreateRichText()
            .AddText("Date: {di:Date}");

        ws.Cell(3, 3)
            .GetRichText()
            .Substring(0, 4)
            .SetFontColor(XLColor.Apricot);
        
        ws.Cell(3, 3)
            .GetRichText()
            .Substring(6, 9)
            .SetBold()
            .SetStrikethrough();

        ws.Cell(3, 2).CreateRichText()
            .AddText("{di:Field1}");

        ws.Cell(3, 2)
            .GetRichText()
            .Substring(0, 4)
            .SetStrikethrough();
        
        ws.Cell(4, 2).CreateRichText()
            .AddText("Field1/Field2: {di:Field1}/{di:Field2}");
        
        ws.Cell(4, 2)
            .GetRichText()
            .Substring(15, 11)
            .SetStrikethrough();
        
        ws.Cell(4, 2)
            .GetRichText()
            .Substring(27, 11)
            .SetUnderline();
        
        ws.Cell(5, 2).CreateRichText()
            .AddText("Sum: {di:parent:Sum} = {di:parent:Sum}");
        
        ws.Cell(5, 2)
            .GetRichText()
            .Substring(0, 21)
            .SetFontColor(XLColor.Green);
        
        ws.Cell(5, 2)
            .GetRichText()
            .Substring(20, 6)
            .SetUnderline();

        ws.Cell(1, 1).CreateRichText()
            .AddText("{di:Name}")
            .SetBold();
        
        ws.Cell(1, 3).CreateRichText()
            .AddText("Name: {di:Name}")
            .SetItalic();
        
        ws.Cell(1, 4).Value = "{di:Name}";
        ws.Cell(3, 1).Value = "{di:Name}";
        ws.Cell(3, 4).Value = "{di:Name}";
        ws.Cell(6, 1).Value = "{di:Name}";
        ws.Cell(6, 3).Value = "{di:Name}";
        ws.Cell(6, 4).Value = "{di:Name}";

        var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"),
            report, report.TemplateProcessor)
        {
            Type = PanelType.Horizontal
        };
        var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
        {
            Parent = parentPanel,
            Type = PanelType.Horizontal
        };
        parentPanel.Children = new[] {childPanel};
        parentPanel.Render();

        Assert.AreEqual(ws.Range(2, 2, 5, 9), parentPanel.ResultRange);
        
        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelRender_WithGrouping_HorizontalPanels_ChildLeft_Test),
            "ParentCellsShiftChildCellsShift_WithRichText"), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void Test_HorizontalPanelsGrouping_ChildLeft_ParentRowShiftChildCellsShift()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var parentRange = ws.Range(2, 2, 5, 3);
        parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

        var child = ws.Range(2, 2, 5, 2);
        child.AddToNamed("ChildRange", XLScope.Worksheet);

        child.Range(2, 1, 4, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        child.Range(2, 1, 4, 1).Style.Border.OutsideBorderColor = XLColor.Red;

        parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

        ws.Cell(2, 3).Value = "{di:Name}";
        ws.Cell(3, 3).Value = "{di:Date}";

        ws.Cell(3, 2).Value = "{di:Field1}";
        ws.Cell(4, 2).Value = "{di:Field2}";
        ws.Cell(5, 2).Value = "{di:parent:Sum}";

        ws.Cell(1, 1).Value = "{di:Name}";
        ws.Cell(1, 3).Value = "{di:Name}";
        ws.Cell(1, 4).Value = "{di:Name}";
        ws.Cell(3, 1).Value = "{di:Name}";
        ws.Cell(3, 4).Value = "{di:Name}";
        ws.Cell(6, 1).Value = "{di:Name}";
        ws.Cell(6, 3).Value = "{di:Name}";
        ws.Cell(6, 4).Value = "{di:Name}";

        var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"),
            report, report.TemplateProcessor)
        {
            Type = PanelType.Horizontal,
            ShiftType = ShiftType.Row
        };
        var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
        {
            Parent = parentPanel,
            Type = PanelType.Horizontal
        };
        parentPanel.Children = new[] {childPanel};
        parentPanel.Render();

        Assert.AreEqual(ws.Range(2, 2, 5, 9), parentPanel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelRender_WithGrouping_HorizontalPanels_ChildLeft_Test),
            "ParentRowShiftChildCellsShift"), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void Test_HorizontalPanelsGrouping_ChildLeft_ParentRowShiftChildRowShift()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var parentRange = ws.Range(2, 2, 5, 3);
        parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

        var child = ws.Range(2, 2, 5, 2);
        child.AddToNamed("ChildRange", XLScope.Worksheet);

        child.Range(2, 1, 4, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        child.Range(2, 1, 4, 1).Style.Border.OutsideBorderColor = XLColor.Red;

        parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

        ws.Cell(2, 3).Value = "{di:Name}";
        ws.Cell(3, 3).Value = "{di:Date}";

        ws.Cell(3, 2).Value = "{di:Field1}";
        ws.Cell(4, 2).Value = "{di:Field2}";
        ws.Cell(5, 2).Value = "{di:parent:Sum}";

        ws.Cell(1, 1).Value = "{di:Name}";
        ws.Cell(1, 3).Value = "{di:Name}";
        ws.Cell(1, 4).Value = "{di:Name}";
        ws.Cell(3, 1).Value = "{di:Name}";
        ws.Cell(3, 4).Value = "{di:Name}";
        ws.Cell(6, 1).Value = "{di:Name}";
        ws.Cell(6, 3).Value = "{di:Name}";
        ws.Cell(6, 4).Value = "{di:Name}";

        var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"),
            report, report.TemplateProcessor)
        {
            Type = PanelType.Horizontal,
            ShiftType = ShiftType.Row
        };
        var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
        {
            Parent = parentPanel,
            Type = PanelType.Horizontal,
            ShiftType = ShiftType.Row
        };
        parentPanel.Children = new[] {childPanel};
        parentPanel.Render();

        Assert.AreEqual(ws.Range(2, 2, 5, 9), parentPanel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelRender_WithGrouping_HorizontalPanels_ChildLeft_Test),
            "ParentRowShiftChildRowShift"), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void Test_HorizontalPanelsGrouping_ChildLeft_ParentNoShiftChildRowShift()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var parentRange = ws.Range(2, 2, 5, 3);
        parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

        var child = ws.Range(2, 2, 5, 2);
        child.AddToNamed("ChildRange", XLScope.Worksheet);

        child.Range(2, 1, 4, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        child.Range(2, 1, 4, 1).Style.Border.OutsideBorderColor = XLColor.Red;

        parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

        ws.Cell(2, 3).Value = "{di:Name}";
        ws.Cell(3, 3).Value = "{di:Date}";

        ws.Cell(3, 2).Value = "{di:Field1}";
        ws.Cell(4, 2).Value = "{di:Field2}";
        ws.Cell(5, 2).Value = "{di:parent:Sum}";

        ws.Cell(1, 1).Value = "{di:Name}";
        ws.Cell(1, 3).Value = "{di:Name}";
        ws.Cell(1, 4).Value = "{di:Name}";
        ws.Cell(3, 1).Value = "{di:Name}";
        ws.Cell(3, 4).Value = "{di:Name}";
        ws.Cell(6, 1).Value = "{di:Name}";
        ws.Cell(6, 3).Value = "{di:Name}";
        ws.Cell(6, 4).Value = "{di:Name}";

        var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"),
            report, report.TemplateProcessor)
        {
            Type = PanelType.Horizontal,
            ShiftType = ShiftType.NoShift
        };
        var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
        {
            Parent = parentPanel,
            Type = PanelType.Horizontal,
            ShiftType = ShiftType.Row
        };
        parentPanel.Children = new[] {childPanel};
        parentPanel.Render();

        Assert.AreEqual(ws.Range(2, 2, 5, 9), parentPanel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelRender_WithGrouping_HorizontalPanels_ChildLeft_Test),
            "ParentNoShiftChildRowShift"), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void Test_HorizontalPanelsGrouping_ChildLeft_ParentCellsShiftChildCellsShift_WithFictitiousColumn()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var parentRange = ws.Range(2, 2, 5, 4);
        parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

        var child = ws.Range(2, 3, 5, 3);
        child.AddToNamed("ChildRange", XLScope.Worksheet);

        child.Range(2, 1, 4, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        child.Range(2, 1, 4, 1).Style.Border.OutsideBorderColor = XLColor.Red;

        parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

        ws.Cell(2, 4).Value = "{di:Name}";
        ws.Cell(3, 4).Value = "{di:Date}";

        ws.Cell(3, 3).Value = "{di:Field1}";
        ws.Cell(4, 3).Value = "{di:Field2}";
        ws.Cell(5, 3).Value = "{di:parent:Sum}";

        ws.Cell(1, 1).Value = "{di:Name}";
        ws.Cell(1, 3).Value = "{di:Name}";
        ws.Cell(1, 5).Value = "{di:Name}";
        ws.Cell(3, 1).Value = "{di:Name}";
        ws.Cell(3, 5).Value = "{di:Name}";
        ws.Cell(6, 1).Value = "{di:Name}";
        ws.Cell(6, 3).Value = "{di:Name}";
        ws.Cell(6, 5).Value = "{di:Name}";

        var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"),
            report, report.TemplateProcessor)
        {
            Type = PanelType.Horizontal
        };
        var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
        {
            Parent = parentPanel,
            Type = PanelType.Horizontal
        };
        parentPanel.Children = new[] {childPanel};
        parentPanel.Render();

        Assert.AreEqual(ws.Range(2, 2, 5, 12), parentPanel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelRender_WithGrouping_HorizontalPanels_ChildLeft_Test),
            "ParentCellsShiftChildCellsShift_WithFictitiousColumn"), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void Test_HorizontalPanelsGrouping_ChildLeft_ParentRowShiftChildCellsShift_WithFictitiousColumn()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var parentRange = ws.Range(2, 2, 5, 4);
        parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

        var child = ws.Range(2, 3, 5, 3);
        child.AddToNamed("ChildRange", XLScope.Worksheet);

        child.Range(2, 1, 4, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        child.Range(2, 1, 4, 1).Style.Border.OutsideBorderColor = XLColor.Red;

        parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

        ws.Cell(2, 4).Value = "{di:Name}";
        ws.Cell(3, 4).Value = "{di:Date}";

        ws.Cell(3, 3).Value = "{di:Field1}";
        ws.Cell(4, 3).Value = "{di:Field2}";
        ws.Cell(5, 3).Value = "{di:parent:Sum}";

        ws.Cell(1, 1).Value = "{di:Name}";
        ws.Cell(1, 3).Value = "{di:Name}";
        ws.Cell(1, 5).Value = "{di:Name}";
        ws.Cell(3, 1).Value = "{di:Name}";
        ws.Cell(3, 5).Value = "{di:Name}";
        ws.Cell(6, 1).Value = "{di:Name}";
        ws.Cell(6, 3).Value = "{di:Name}";
        ws.Cell(6, 5).Value = "{di:Name}";

        var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"),
            report, report.TemplateProcessor)
        {
            Type = PanelType.Horizontal,
            ShiftType = ShiftType.Row
        };
        var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
        {
            Parent = parentPanel,
            Type = PanelType.Horizontal
        };
        parentPanel.Children = new[] {childPanel};
        parentPanel.Render();

        Assert.AreEqual(ws.Range(2, 2, 5, 12), parentPanel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelRender_WithGrouping_HorizontalPanels_ChildLeft_Test),
            "ParentRowShiftChildCellsShift_WithFictitiousColumn"), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void Test_HorizontalPanelsGrouping_ChildLeft_ParentRowShiftChildRowShift_WithFictitiousColumn()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var parentRange = ws.Range(2, 2, 5, 4);
        parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

        var child = ws.Range(2, 3, 5, 3);
        child.AddToNamed("ChildRange", XLScope.Worksheet);

        child.Range(2, 1, 4, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        child.Range(2, 1, 4, 1).Style.Border.OutsideBorderColor = XLColor.Red;

        parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

        ws.Cell(2, 4).Value = "{di:Name}";
        ws.Cell(3, 4).Value = "{di:Date}";

        ws.Cell(3, 3).Value = "{di:Field1}";
        ws.Cell(4, 3).Value = "{di:Field2}";
        ws.Cell(5, 3).Value = "{di:parent:Sum}";

        ws.Cell(1, 1).Value = "{di:Name}";
        ws.Cell(1, 3).Value = "{di:Name}";
        ws.Cell(1, 5).Value = "{di:Name}";
        ws.Cell(3, 1).Value = "{di:Name}";
        ws.Cell(3, 5).Value = "{di:Name}";
        ws.Cell(6, 1).Value = "{di:Name}";
        ws.Cell(6, 3).Value = "{di:Name}";
        ws.Cell(6, 5).Value = "{di:Name}";

        var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"),
            report, report.TemplateProcessor)
        {
            Type = PanelType.Horizontal,
            ShiftType = ShiftType.Row
        };
        var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
        {
            Parent = parentPanel,
            Type = PanelType.Horizontal,
            ShiftType = ShiftType.Row
        };
        parentPanel.Children = new[] {childPanel};
        parentPanel.Render();

        Assert.AreEqual(ws.Range(2, 2, 5, 12), parentPanel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelRender_WithGrouping_HorizontalPanels_ChildLeft_Test),
            "ParentRowShiftChildRowShift_WithFictitiousColumn"), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void Test_HorizontalPanelsGrouping_ChildLeft_ParentNoShiftChildCellsShift_WithFictitiousColumn()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var parentRange = ws.Range(2, 2, 5, 4);
        parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

        var child = ws.Range(2, 3, 5, 3);
        child.AddToNamed("ChildRange", XLScope.Worksheet);

        child.Range(2, 1, 4, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        child.Range(2, 1, 4, 1).Style.Border.OutsideBorderColor = XLColor.Red;

        parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

        ws.Cell(2, 4).Value = "{di:Name}";
        ws.Cell(3, 4).Value = "{di:Date}";

        ws.Cell(3, 3).Value = "{di:Field1}";
        ws.Cell(4, 3).Value = "{di:Field2}";
        ws.Cell(5, 3).Value = "{di:parent:Sum}";

        ws.Cell(1, 1).Value = "{di:Name}";
        ws.Cell(1, 3).Value = "{di:Name}";
        ws.Cell(1, 5).Value = "{di:Name}";
        ws.Cell(3, 1).Value = "{di:Name}";
        ws.Cell(3, 5).Value = "{di:Name}";
        ws.Cell(6, 1).Value = "{di:Name}";
        ws.Cell(6, 3).Value = "{di:Name}";
        ws.Cell(6, 5).Value = "{di:Name}";

        var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"),
            report, report.TemplateProcessor)
        {
            Type = PanelType.Horizontal,
            ShiftType = ShiftType.NoShift
        };
        var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
        {
            Parent = parentPanel,
            Type = PanelType.Horizontal
        };
        parentPanel.Children = new[] {childPanel};
        parentPanel.Render();

        Assert.AreEqual(ws.Range(2, 2, 5, 12), parentPanel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelRender_WithGrouping_HorizontalPanels_ChildLeft_Test),
            "ParentNoShiftChildCellsShift_WithFictitiousColumn"), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void
        Test_HorizontalPanelsGrouping_ChildLeft_ParentCellsShiftChildCellsShift_WithFictitiousColumnWhichDeleteAfterRender()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var parentRange = ws.Range(2, 2, 5, 4);
        parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

        var child = ws.Range(2, 3, 5, 3);
        child.AddToNamed("ChildRange", XLScope.Worksheet);

        child.Range(2, 1, 4, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        child.Range(2, 1, 4, 1).Style.Border.OutsideBorderColor = XLColor.Red;

        parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

        ws.Cell(2, 4).Value = "{di:Name}";
        ws.Cell(3, 4).Value = "{di:Date}";

        ws.Cell(3, 3).Value = "{di:Field1}";
        ws.Cell(4, 3).Value = "{di:Field2}";
        ws.Cell(5, 3).Value = "{di:parent:Sum}";

        ws.Cell(1, 1).Value = "{di:Name}";
        ws.Cell(1, 3).Value = "{di:Name}";
        ws.Cell(1, 5).Value = "{di:Name}";
        ws.Cell(3, 1).Value = "{di:Name}";
        ws.Cell(3, 5).Value = "{di:Name}";
        ws.Cell(6, 1).Value = "{di:Name}";
        ws.Cell(6, 3).Value = "{di:Name}";
        ws.Cell(6, 5).Value = "{di:Name}";

        var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"),
            report, report.TemplateProcessor)
        {
            Type = PanelType.Horizontal,
            AfterDataItemRenderMethodName = "AfterRenderParentDataSourcePanelChildLeft"
        };
        var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)",
            ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
        {
            Parent = parentPanel,
            Type = PanelType.Horizontal
        };
        parentPanel.Children = new[] {childPanel};
        parentPanel.Render();

        Assert.AreEqual(ws.Range(2, 2, 5, 9), parentPanel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelRender_WithGrouping_HorizontalPanels_ChildLeft_Test),
            "ParentCellsShiftChildCellsShift_WithFictitiousColumnWhichDeleteAfterRender"), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }
}