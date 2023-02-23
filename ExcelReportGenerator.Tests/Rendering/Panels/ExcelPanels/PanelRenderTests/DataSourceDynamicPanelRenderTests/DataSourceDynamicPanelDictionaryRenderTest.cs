using ClosedXML.Excel;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourceDynamicPanelRenderTests;

public class DataSourceDynamicPanelDictionaryRenderTest
{
    [Test]
    public void TestRenderDictionary()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range1 = ws.Range(2, 2, 4, 2);
        range1.AddToNamed("TestRange", XLScope.Worksheet);

        ws.Cell(2, 2).Value = "{Headers}";
        ws.Cell(3, 2).Value = "{Data}";
        ws.Cell(4, 2).Value = "{Totals}";

        var range2 = ws.Range(7, 2, 9, 2);
        range2.AddToNamed("TestRange2", XLScope.Worksheet);

        ws.Cell(7, 2).Value = "{Headers}";
        ws.Cell(8, 2).Value = "{Data}";
        ws.Cell(9, 2).Value = "{Totals}";

        var data1 = new DataProvider().GetDictionaryEnumerable().First();
        var panel1 =
            new ExcelDataSourceDynamicPanel(data1, ws.NamedRange("TestRange"), report, report.TemplateProcessor);
        panel1.Render();

        Assert.AreEqual(ws.Range(2, 2, 6, 3), panel1.ResultRange);

        var data2 = new DataProvider().GetDictionaryEnumerable().First()
            .Select(x => new KeyValuePair<string, object>(x.Key, x.Value));
        var panel2 =
            new ExcelDataSourceDynamicPanel(data2, ws.NamedRange("TestRange2"), report, report.TemplateProcessor);
        panel2.Render();

        Assert.AreEqual(ws.Range(9, 2, 13, 3), panel2.ResultRange);

        ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(
            nameof(DataSourceDynamicPanelDictionaryRenderTest),
            nameof(TestRenderDictionary)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void TestRenderEmptyDictionary()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range1 = ws.Range(2, 2, 4, 2);
        range1.AddToNamed("TestRange", XLScope.Worksheet);

        ws.Cell(2, 2).Value = "{Headers}";
        ws.Cell(3, 2).Value = "{Data}";
        ws.Cell(4, 2).Value = "{Totals}";

        var range2 = ws.Range(7, 2, 9, 2);
        range2.AddToNamed("TestRange2", XLScope.Worksheet);

        ws.Cell(7, 2).Value = "{Headers}";
        ws.Cell(8, 2).Value = "{Data}";
        ws.Cell(9, 2).Value = "{Totals}";

        IDictionary<string, object> data1 = new Dictionary<string, object>();
        var panel1 =
            new ExcelDataSourceDynamicPanel(data1, ws.NamedRange("TestRange"), report, report.TemplateProcessor);
        panel1.Render();

        Assert.AreEqual(ws.Range(2, 2, 3, 3), panel1.ResultRange);

        IEnumerable<KeyValuePair<string, object>> data2 = new List<KeyValuePair<string, object>>();
        var panel2 =
            new ExcelDataSourceDynamicPanel(data2, ws.NamedRange("TestRange2"), report, report.TemplateProcessor);
        panel2.Render();

        Assert.AreEqual(ws.Range(6, 2, 7, 3), panel2.ResultRange);

        Assert.AreEqual(4, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
        Assert.AreEqual("Key", ws.Cell(2, 2).Value);
        Assert.AreEqual("Value", ws.Cell(2, 3).Value);
        Assert.AreEqual("Key", ws.Cell(6, 2).Value);
        Assert.AreEqual("Value", ws.Cell(6, 3).Value);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void TestRenderDictionaryEnumerable()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range1 = ws.Range(2, 2, 4, 2);
        range1.AddToNamed("TestRange1", XLScope.Worksheet);

        ws.Cell(2, 2).Value = "{Headers}";
        ws.Cell(3, 2).Value = "{Data}";
        ws.Cell(4, 2).Value = "{Totals}";

        var panel1 = new ExcelDataSourceDynamicPanel("m:DataProvider:GetDictionaryEnumerable()",
            ws.NamedRange("TestRange1"), report, report.TemplateProcessor);
        panel1.Render();

        Assert.AreEqual(ws.Range(2, 2, 6, 4), panel1.ResultRange);

        var dictWithDecimalValues = new List<IDictionary<string, decimal>>
        {
            new Dictionary<string, decimal> {["Value"] = 25.7m},
            new Dictionary<string, decimal> {["Value"] = 250.7m},
            new Dictionary<string, decimal> {["Value"] = 2500.7m}
        };

        var range2 = ws.Range(7, 2, 9, 2);
        range2.AddToNamed("TestRange2", XLScope.Worksheet);

        ws.Cell(7, 2).Value = "{Headers}";
        ws.Cell(8, 2).Value = "{Data}";
        ws.Cell(9, 2).Value = "{Totals}";

        var panel2 = new ExcelDataSourceDynamicPanel(dictWithDecimalValues, ws.NamedRange("TestRange2"), report,
            report.TemplateProcessor);
        panel2.Render();

        Assert.AreEqual(ws.Range(7, 2, 11, 2), panel2.ResultRange);

        ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(
            nameof(DataSourceDynamicPanelDictionaryRenderTest),
            nameof(TestRenderDictionaryEnumerable)), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    [Test]
    public void TestRenderEmptyDictionaryEnumerable()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range = ws.Range(2, 2, 4, 2);
        range.AddToNamed("TestRange1", XLScope.Worksheet);

        ws.Cell(2, 2).Value = "{Headers}";
        ws.Cell(3, 2).Value = "{Data}";
        ws.Cell(4, 2).Value = "{Totals}";

        var panel = new ExcelDataSourceDynamicPanel(new List<IDictionary<string, decimal>>(),
            ws.NamedRange("TestRange1"), report, report.TemplateProcessor);
        panel.Render();

        Assert.IsNull(panel.ResultRange);

        Assert.AreEqual(0, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());

        //report.Workbook.SaveAs("test.xlsx");
    }
}