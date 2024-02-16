﻿using ClosedXML.Excel;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;
using Assert = NUnit.Framework.Legacy.ClassicAssert;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourcePanelRenderTests;

public class DataSourcePanelEmptyIEnumerableRenderTest
{
    [Test]
    public void TestRenderEmptyIEnumerableVerticalCellsShift()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range = ws.Range(2, 2, 3, 5);
        range.AddToNamed("TestRange", XLScope.Worksheet);

        range.Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
        range.Style.Border.SetRightBorder(XLBorderStyleValues.Thin);
        range.Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
        range.Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

        ws.Cell(4, 3).Style.Border.SetTopBorder(XLBorderStyleValues.Thin);

        ws.Cell(2, 2).Value = "{di:Name}";
        ws.Cell(2, 3).Value = "{di:Date}";
        ws.Cell(2, 4).Value = "{di:Sum}";
        ws.Cell(2, 5).Value = "{di:Contacts}";
        ws.Cell(3, 2).Value = "{di:Contacts.Phone}";
        ws.Cell(3, 3).Value = "{di:Contacts.Fax}";
        ws.Cell(3, 4).Value = "{p:StrParam}";

        ws.Cell(1, 1).Value = "{di:Name}";
        ws.Cell(4, 1).Value = "{di:Name}";
        ws.Cell(1, 6).Value = "{di:Name}";
        ws.Cell(4, 6).Value = "{di:Name}";
        ws.Cell(3, 1).Value = "{di:Name}";
        ws.Cell(3, 6).Value = "{di:Name}";
        ws.Cell(1, 4).Value = "{di:Name}";
        ws.Cell(4, 4).Value = "{di:Name}";

        var panel = new ExcelDataSourcePanel("m:DataProvider:GetEmptyIEnumerable()", ws.NamedRange("TestRange"), report,
            report.TemplateProcessor);
        panel.Render();

        Assert.IsNull(panel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelEmptyIEnumerableRenderTest),
            nameof(TestRenderEmptyIEnumerableVerticalCellsShift)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void TestRenderEmptyIEnumerableVerticalRowsShift()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range = ws.Range(2, 2, 3, 5);
        range.AddToNamed("TestRange", XLScope.Worksheet);

        range.Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
        range.Style.Border.SetRightBorder(XLBorderStyleValues.Thin);
        range.Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
        range.Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

        ws.Cell(4, 3).Style.Border.SetTopBorder(XLBorderStyleValues.Thin);

        ws.Cell(2, 2).Value = "{di:Name}";
        ws.Cell(2, 3).Value = "{di:Date}";
        ws.Cell(2, 4).Value = "{di:Sum}";
        ws.Cell(2, 5).Value = "{di:Contacts}";
        ws.Cell(3, 2).Value = "{di:Contacts.Phone}";
        ws.Cell(3, 3).Value = "{di:Contacts.Fax}";
        ws.Cell(3, 4).Value = "{p:StrParam}";

        ws.Cell(1, 1).Value = "{di:Name}";
        ws.Cell(4, 1).Value = "{di:Name}";
        ws.Cell(1, 6).Value = "{di:Name}";
        ws.Cell(4, 6).Value = "{di:Name}";
        ws.Cell(3, 1).Value = "{di:Name}";
        ws.Cell(3, 6).Value = "{di:Name}";
        ws.Cell(1, 4).Value = "{di:Name}";
        ws.Cell(4, 4).Value = "{di:Name}";

        var panel = new ExcelDataSourcePanel("m:DataProvider:GetEmptyIEnumerable()", ws.NamedRange("TestRange"), report,
            report.TemplateProcessor)
        {
            ShiftType = ShiftType.Row
        };
        panel.Render();

        Assert.IsNull(panel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelEmptyIEnumerableRenderTest),
            nameof(TestRenderEmptyIEnumerableVerticalRowsShift)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void TestRenderEmptyIEnumerableVerticalNoShift()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range = ws.Range(2, 2, 3, 5);
        range.AddToNamed("TestRange", XLScope.Worksheet);

        range.Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
        range.Style.Border.SetRightBorder(XLBorderStyleValues.Thin);
        range.Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
        range.Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

        ws.Cell(4, 3).Style.Border.SetTopBorder(XLBorderStyleValues.Thin);

        ws.Cell(2, 2).Value = "{di:Name}";
        ws.Cell(2, 3).Value = "{di:Date}";
        ws.Cell(2, 4).Value = "{di:Sum}";
        ws.Cell(2, 5).Value = "{di:Contacts}";
        ws.Cell(3, 2).Value = "{di:Contacts.Phone}";
        ws.Cell(3, 3).Value = "{di:Contacts.Fax}";
        ws.Cell(3, 4).Value = "{p:StrParam}";

        ws.Cell(1, 1).Value = "{di:Name}";
        ws.Cell(4, 1).Value = "{di:Name}";
        ws.Cell(1, 6).Value = "{di:Name}";
        ws.Cell(4, 6).Value = "{di:Name}";
        ws.Cell(3, 1).Value = "{di:Name}";
        ws.Cell(3, 6).Value = "{di:Name}";
        ws.Cell(1, 4).Value = "{di:Name}";
        ws.Cell(4, 4).Value = "{di:Name}";

        var panel = new ExcelDataSourcePanel("m:DataProvider:GetEmptyIEnumerable()", ws.NamedRange("TestRange"), report,
            report.TemplateProcessor)
        {
            ShiftType = ShiftType.NoShift
        };
        panel.Render();

        Assert.IsNull(panel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelEmptyIEnumerableRenderTest),
            nameof(TestRenderEmptyIEnumerableVerticalNoShift)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void TestRenderEmptyIEnumerableHorizontalCellsShift()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range = ws.Range(2, 2, 3, 5);
        range.AddToNamed("TestRange", XLScope.Worksheet);

        range.Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
        range.Style.Border.SetRightBorder(XLBorderStyleValues.Thin);
        range.Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
        range.Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

        ws.Cell(2, 6).Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

        ws.Cell(2, 2).Value = "{di:Name}";
        ws.Cell(2, 3).Value = "{di:Date}";
        ws.Cell(2, 4).Value = "{di:Sum}";
        ws.Cell(2, 5).Value = "{di:Contacts}";
        ws.Cell(3, 2).Value = "{di:Contacts.Phone}";
        ws.Cell(3, 3).Value = "{di:Contacts.Fax}";
        ws.Cell(3, 4).Value = "{p:StrParam}";

        ws.Cell(1, 1).Value = "{di:Name}";
        ws.Cell(4, 1).Value = "{di:Name}";
        ws.Cell(1, 6).Value = "{di:Name}";
        ws.Cell(4, 6).Value = "{di:Name}";
        ws.Cell(3, 1).Value = "{di:Name}";
        ws.Cell(3, 6).Value = "{di:Name}";
        ws.Cell(1, 4).Value = "{di:Name}";
        ws.Cell(4, 4).Value = "{di:Name}";

        var panel = new ExcelDataSourcePanel("m:DataProvider:GetEmptyIEnumerable()", ws.NamedRange("TestRange"), report,
            report.TemplateProcessor)
        {
            Type = PanelType.Horizontal
        };
        panel.Render();

        Assert.IsNull(panel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelEmptyIEnumerableRenderTest),
            nameof(TestRenderEmptyIEnumerableHorizontalCellsShift)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void TestRenderEmptyIEnumerableHorizontalRowsShift()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range = ws.Range(2, 2, 3, 5);
        range.AddToNamed("TestRange", XLScope.Worksheet);

        range.Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
        range.Style.Border.SetRightBorder(XLBorderStyleValues.Thin);
        range.Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
        range.Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

        ws.Cell(2, 6).Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

        ws.Cell(2, 2).Value = "{di:Name}";
        ws.Cell(2, 3).Value = "{di:Date}";
        ws.Cell(2, 4).Value = "{di:Sum}";
        ws.Cell(2, 5).Value = "{di:Contacts}";
        ws.Cell(3, 2).Value = "{di:Contacts.Phone}";
        ws.Cell(3, 3).Value = "{di:Contacts.Fax}";
        ws.Cell(3, 4).Value = "{p:StrParam}";

        ws.Cell(1, 1).Value = "{di:Name}";
        ws.Cell(4, 1).Value = "{di:Name}";
        ws.Cell(1, 6).Value = "{di:Name}";
        ws.Cell(4, 6).Value = "{di:Name}";
        ws.Cell(3, 1).Value = "{di:Name}";
        ws.Cell(3, 6).Value = "{di:Name}";
        ws.Cell(1, 4).Value = "{di:Name}";
        ws.Cell(4, 4).Value = "{di:Name}";

        var panel = new ExcelDataSourcePanel("m:DataProvider:GetEmptyIEnumerable()", ws.NamedRange("TestRange"), report,
            report.TemplateProcessor)
        {
            Type = PanelType.Horizontal,
            ShiftType = ShiftType.Row
        };
        panel.Render();

        Assert.IsNull(panel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelEmptyIEnumerableRenderTest),
            nameof(TestRenderEmptyIEnumerableHorizontalRowsShift)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void TestRenderEmptyIEnumerableHorizontalNoShift()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range = ws.Range(2, 2, 3, 5);
        range.AddToNamed("TestRange", XLScope.Worksheet);

        range.Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
        range.Style.Border.SetRightBorder(XLBorderStyleValues.Thin);
        range.Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
        range.Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

        ws.Cell(2, 6).Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

        ws.Cell(2, 2).Value = "{di:Name}";
        ws.Cell(2, 3).Value = "{di:Date}";
        ws.Cell(2, 4).Value = "{di:Sum}";
        ws.Cell(2, 5).Value = "{di:Contacts}";
        ws.Cell(3, 2).Value = "{di:Contacts.Phone}";
        ws.Cell(3, 3).Value = "{di:Contacts.Fax}";
        ws.Cell(3, 4).Value = "{p:StrParam}";

        ws.Cell(1, 1).Value = "{di:Name}";
        ws.Cell(4, 1).Value = "{di:Name}";
        ws.Cell(1, 6).Value = "{di:Name}";
        ws.Cell(4, 6).Value = "{di:Name}";
        ws.Cell(3, 1).Value = "{di:Name}";
        ws.Cell(3, 6).Value = "{di:Name}";
        ws.Cell(1, 4).Value = "{di:Name}";
        ws.Cell(4, 4).Value = "{di:Name}";

        var panel = new ExcelDataSourcePanel("m:DataProvider:GetEmptyIEnumerable()", ws.NamedRange("TestRange"), report,
            report.TemplateProcessor)
        {
            Type = PanelType.Horizontal,
            ShiftType = ShiftType.NoShift
        };
        panel.Render();

        Assert.IsNull(panel.ResultRange);

        ExcelAssert.AreWorkbooksContentEqual(TestHelper.GetExpectedWorkbook(
            nameof(DataSourcePanelEmptyIEnumerableRenderTest),
            nameof(TestRenderEmptyIEnumerableHorizontalNoShift)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }
}