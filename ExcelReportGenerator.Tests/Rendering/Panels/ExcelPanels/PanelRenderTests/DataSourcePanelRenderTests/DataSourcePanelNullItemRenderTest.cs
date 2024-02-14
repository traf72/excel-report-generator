using ClosedXML.Excel;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;
using Assert = NUnit.Framework.Legacy.ClassicAssert;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourcePanelRenderTests;

public class DataSourcePanelNullItemRenderTest
{
    [Test]
    public void TestRenderNullItemVerticalCellsShift()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range = ws.Range(2, 2, 3, 5);
        range.AddToNamed("TestRange", XLScope.Worksheet);

        ws.Cell(2, 2).Value = "{di:Name}";
        ws.Cell(2, 3).Value = "{di:Date}";
        ws.Cell(2, 4).Value = "{m:Multiply(di:Sum, 5)}";
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

        var panel = new ExcelDataSourcePanel("m:DataProvider:GetNullItem()", ws.NamedRange("TestRange"), report,
            report.TemplateProcessor);
        panel.Render();

        Assert.IsNull(panel.ResultRange);

        ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelNullItemRenderTest),
            nameof(TestRenderNullItemVerticalCellsShift)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void TestRenderNullItemVerticalRowShift()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range = ws.Range(2, 2, 3, 5);
        range.AddToNamed("TestRange", XLScope.Worksheet);

        ws.Cell(2, 2).Value = "{di:Name}";
        ws.Cell(2, 3).Value = "{di:Date}";
        ws.Cell(2, 4).Value = "{m:Multiply(di:Sum, 5)}";
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

        var panel = new ExcelDataSourcePanel("m:DataProvider:GetNullItem()", ws.NamedRange("TestRange"), report,
            report.TemplateProcessor)
        {
            ShiftType = ShiftType.Row
        };
        panel.Render();

        Assert.IsNull(panel.ResultRange);

        ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelNullItemRenderTest),
            nameof(TestRenderNullItemVerticalRowShift)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void TestRenderNullItemVerticalNoShift()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range = ws.Range(2, 2, 3, 5);
        range.AddToNamed("TestRange", XLScope.Worksheet);

        ws.Cell(2, 2).Value = "{di:Name}";
        ws.Cell(2, 3).Value = "{di:Date}";
        ws.Cell(2, 4).Value = "{m:Multiply(di:Sum, 5)}";
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

        var panel = new ExcelDataSourcePanel("m:DataProvider:GetNullItem()", ws.NamedRange("TestRange"), report,
            report.TemplateProcessor)
        {
            ShiftType = ShiftType.NoShift
        };
        panel.Render();

        Assert.IsNull(panel.ResultRange);

        ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelNullItemRenderTest),
            nameof(TestRenderNullItemVerticalNoShift)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void TestRenderNullItemHorizontalCellsShift()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range = ws.Range(2, 2, 3, 5);
        range.AddToNamed("TestRange", XLScope.Worksheet);

        ws.Cell(2, 2).Value = "{di:Name}";
        ws.Cell(2, 3).Value = "{di:Date}";
        ws.Cell(2, 4).Value = "{m:Multiply(di:Sum, 5)}";
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

        var panel = new ExcelDataSourcePanel("m:DataProvider:GetNullItem()", ws.NamedRange("TestRange"), report,
            report.TemplateProcessor)
        {
            Type = PanelType.Horizontal
        };
        panel.Render();

        Assert.IsNull(panel.ResultRange);

        ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelNullItemRenderTest),
            nameof(TestRenderNullItemHorizontalCellsShift)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void TestRenderNullItemHorizontalRowShift()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range = ws.Range(2, 2, 3, 5);
        range.AddToNamed("TestRange", XLScope.Worksheet);

        ws.Cell(2, 2).Value = "{di:Name}";
        ws.Cell(2, 3).Value = "{di:Date}";
        ws.Cell(2, 4).Value = "{m:Multiply(di:Sum, 5)}";
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

        var panel = new ExcelDataSourcePanel("m:DataProvider:GetNullItem()", ws.NamedRange("TestRange"), report,
            report.TemplateProcessor)
        {
            Type = PanelType.Horizontal,
            ShiftType = ShiftType.Row
        };
        panel.Render();

        Assert.IsNull(panel.ResultRange);

        ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelNullItemRenderTest),
            nameof(TestRenderNullItemHorizontalRowShift)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void TestRenderNullItemHorizontalNoShift()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range = ws.Range(2, 2, 3, 5);
        range.AddToNamed("TestRange", XLScope.Worksheet);

        ws.Cell(2, 2).Value = "{di:Name}";
        ws.Cell(2, 3).Value = "{di:Date}";
        ws.Cell(2, 4).Value = "{m:Multiply(di:Sum, 5)}";
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

        var panel = new ExcelDataSourcePanel("m:DataProvider:GetNullItem()", ws.NamedRange("TestRange"), report,
            report.TemplateProcessor)
        {
            Type = PanelType.Horizontal,
            ShiftType = ShiftType.NoShift
        };
        panel.Render();

        Assert.IsNull(panel.ResultRange);

        ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelNullItemRenderTest),
            nameof(TestRenderNullItemHorizontalNoShift)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }
}