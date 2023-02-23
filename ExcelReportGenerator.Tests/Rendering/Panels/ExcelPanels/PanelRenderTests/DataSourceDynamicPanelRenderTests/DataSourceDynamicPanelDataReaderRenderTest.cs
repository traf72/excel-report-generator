using ClosedXML.Excel;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourceDynamicPanelRenderTests;

public class DataSourceDynamicPanelDataReaderRenderTest
{
    [Test]
    public void TestRenderDataReader()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range = ws.Range(2, 2, 7, 2);
        range.AddToNamed("TestRange", XLScope.Worksheet);

        ws.Cell(2, 2).Value = "{Headers}";
        ws.Cell(2, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
        ws.Cell(2, 2).Style.Border.OutsideBorderColor = XLColor.Red;
        ws.Cell(2, 2).Style.Font.Bold = true;

        ws.Cell(3, 2).Value = "{Numbers}";
        ws.Cell(3, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.Cell(3, 2).Style.Border.OutsideBorderColor = XLColor.Black;
        ws.Cell(3, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        ws.Cell(4, 2).Value = "{Data}";
        ws.Cell(4, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.Cell(4, 2).Style.Border.OutsideBorderColor = XLColor.Black;

        ws.Cell(5, 2).Value = "{Totals}";
        ws.Cell(5, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Dotted;
        ws.Cell(5, 2).Style.Border.OutsideBorderColor = XLColor.Green;

        ws.Cell(7, 2).FormulaA1 = "=COLUMN()";
        ws.Cell(7, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Dashed;
        ws.Cell(7, 2).Style.Border.OutsideBorderColor = XLColor.Blue;

        var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetAllCustomersDataReader()",
            ws.NamedRange("TestRange"), report, report.TemplateProcessor);
        panel.Render();

        Assert.AreEqual(ws.Range(2, 2, 9, 7), panel.ResultRange);

        ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(
            nameof(DataSourceDynamicPanelDataReaderRenderTest),
            nameof(TestRenderDataReader)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void TestRenderDataReader_HorizontalPanel()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range = ws.Range(2, 2, 2, 5);
        range.AddToNamed("TestRange", XLScope.Worksheet);

        ws.Cell(2, 2).Value = "{Headers}";
        ws.Cell(2, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
        ws.Cell(2, 2).Style.Border.OutsideBorderColor = XLColor.Red;
        ws.Cell(2, 2).Style.Font.Bold = true;

        ws.Cell(2, 3).Value = "{Numbers(5)}";
        ws.Cell(2, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.Cell(2, 3).Style.Border.OutsideBorderColor = XLColor.Black;
        ws.Cell(2, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        ws.Cell(2, 4).Value = "{Data}";
        ws.Cell(2, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.Cell(2, 4).Style.Border.OutsideBorderColor = XLColor.Black;

        ws.Cell(2, 5).Value = "{Totals}";
        ws.Cell(2, 5).Style.Border.OutsideBorder = XLBorderStyleValues.Dotted;
        ws.Cell(2, 5).Style.Border.OutsideBorderColor = XLColor.Green;

        var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetAllCustomersDataReader()",
            ws.NamedRange("TestRange"), report, report.TemplateProcessor)
        {
            Type = PanelType.Horizontal
        };
        panel.Render();

        Assert.AreEqual(ws.Range(2, 2, 7, 7), panel.ResultRange);

        ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(
            nameof(DataSourceDynamicPanelDataReaderRenderTest),
            nameof(TestRenderDataReader_HorizontalPanel)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void TestRenderEmptyDataReader()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range = ws.Range(2, 2, 4, 2);
        range.AddToNamed("TestRange", XLScope.Worksheet);

        ws.Cell(2, 2).Value = "{Headers}";
        ws.Cell(2, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
        ws.Cell(2, 2).Style.Border.OutsideBorderColor = XLColor.Red;
        ws.Cell(2, 2).Style.Font.Bold = true;

        ws.Cell(3, 2).Value = "{Data}";
        ws.Cell(3, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.Cell(3, 2).Style.Border.OutsideBorderColor = XLColor.Black;

        ws.Cell(4, 2).Value = "{Totals}";
        ws.Cell(4, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Dotted;
        ws.Cell(4, 2).Style.Border.OutsideBorderColor = XLColor.Green;

        var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetEmptyDataReader()", ws.NamedRange("TestRange"),
            report, report.TemplateProcessor);
        panel.Render();

        Assert.AreEqual(ws.Range(2, 2, 3, 7), panel.ResultRange);

        ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(
            nameof(DataSourceDynamicPanelDataReaderRenderTest),
            nameof(TestRenderEmptyDataReader)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }
}