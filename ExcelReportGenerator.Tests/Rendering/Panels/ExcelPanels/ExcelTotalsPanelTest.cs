using System.Collections;
using System.Data;
using System.Reflection;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using ExcelReportGenerator.Enumerators;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Exceptions;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using ExcelReportGenerator.Tests.CustomAsserts;
using ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests;
using Microsoft.CSharp.RuntimeBinder;
using Assert = NUnit.Framework.Legacy.ClassicAssert;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels;

public class ExcelTotalsPanelTest
{
    [Test]
    public void TestCopyIfDataSourceTemplateIsSet()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");
        var excelReport = Substitute.For<object>();
        var templateProcessor = Substitute.For<ITemplateProcessor>();

        var range = ws.Range(1, 1, 2, 4);
        range.AddToNamed("DataPanel", XLScope.Worksheet);
        var namedRange = ws.NamedRange("DataPanel");

        var panel = new ExcelTotalsPanel("m:GetData()", namedRange, excelReport, templateProcessor)
        {
            RenderPriority = 10,
            Type = PanelType.Horizontal,
            ShiftType = ShiftType.NoShift,
            BeforeRenderMethodName = "BeforeRenderMethod",
            AfterRenderMethodName = "AfterRenderMethod",
            BeforeDataItemRenderMethodName = "BeforeDataItemRenderMethodName",
            AfterDataItemRenderMethodName = "AfterDataItemRenderMethodName",
            GroupBy = "2,4"
        };

        var copiedPanel = (ExcelTotalsPanel) panel.Copy(ws.Cell(5, 5));

        Assert.AreSame(excelReport,
            copiedPanel.GetType().GetField("_report", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(copiedPanel));
        Assert.AreSame(templateProcessor,
            copiedPanel.GetType().GetField("_templateProcessor", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(copiedPanel));
        Assert.IsNull(copiedPanel.GetType().GetField("_data", BindingFlags.Instance | BindingFlags.NonPublic)
            .GetValue(copiedPanel));
        Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Range.FirstCell());
        Assert.AreEqual(ws.Cell(6, 8), copiedPanel.Range.LastCell());
        Assert.AreEqual(10, copiedPanel.RenderPriority);
        Assert.AreEqual(PanelType.Horizontal, copiedPanel.Type);
        Assert.AreEqual(ShiftType.NoShift, copiedPanel.ShiftType);
        Assert.AreEqual("BeforeRenderMethod", copiedPanel.BeforeRenderMethodName);
        Assert.AreEqual("AfterRenderMethod", copiedPanel.AfterRenderMethodName);
        Assert.AreEqual("BeforeDataItemRenderMethodName", copiedPanel.BeforeDataItemRenderMethodName);
        Assert.AreEqual("AfterDataItemRenderMethodName", copiedPanel.AfterDataItemRenderMethodName);
        Assert.AreEqual("2,4", copiedPanel.GroupBy);
        Assert.IsNull(copiedPanel.Parent);

        //wb.SaveAs("test.xlsx");
    }

    [Test]
    public void TestCopyIfDataIsSet()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");
        var excelReport = Substitute.For<object>();
        var templateProcessor = Substitute.For<ITemplateProcessor>();

        var range = ws.Range(1, 1, 2, 4);
        range.AddToNamed("DataPanel", XLScope.Worksheet);
        var namedRange = ws.NamedRange("DataPanel");

        object[] data = {1, "One"};
        var panel = new ExcelTotalsPanel(data, namedRange, excelReport, templateProcessor)
        {
            RenderPriority = 10,
            Type = PanelType.Horizontal,
            ShiftType = ShiftType.NoShift,
            BeforeRenderMethodName = "BeforeRenderMethod",
            AfterRenderMethodName = "AfterRenderMethod",
            BeforeDataItemRenderMethodName = "BeforeDataItemRenderMethodName",
            AfterDataItemRenderMethodName = "AfterDataItemRenderMethodName",
            GroupBy = "2,4"
        };

        var copiedPanel = (ExcelTotalsPanel) panel.Copy(ws.Cell(5, 5));

        Assert.AreSame(excelReport,
            copiedPanel.GetType().GetField("_report", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(copiedPanel));
        Assert.AreSame(templateProcessor,
            copiedPanel.GetType().GetField("_templateProcessor", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(copiedPanel));
        Assert.IsNull(copiedPanel.GetType()
            .GetField("_dataSourceTemplate", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(copiedPanel));
        Assert.AreSame(data,
            copiedPanel.GetType().GetField("_data", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(copiedPanel));
        Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Range.FirstCell());
        Assert.AreEqual(ws.Cell(6, 8), copiedPanel.Range.LastCell());
        Assert.AreEqual(10, copiedPanel.RenderPriority);
        Assert.AreEqual(PanelType.Horizontal, copiedPanel.Type);
        Assert.AreEqual(ShiftType.NoShift, copiedPanel.ShiftType);
        Assert.AreEqual("BeforeRenderMethod", copiedPanel.BeforeRenderMethodName);
        Assert.AreEqual("AfterRenderMethod", copiedPanel.AfterRenderMethodName);
        Assert.AreEqual("BeforeDataItemRenderMethodName", copiedPanel.BeforeDataItemRenderMethodName);
        Assert.AreEqual("AfterDataItemRenderMethodName", copiedPanel.AfterDataItemRenderMethodName);
        Assert.AreEqual("2,4", copiedPanel.GroupBy);
        Assert.IsNull(copiedPanel.Parent);

        //wb.SaveAs("test.xlsx");
    }

    [Test]
    public void TestDoAggregation()
    {
        var dataTable = new DataTable();
        dataTable.Columns.Add(new DataColumn("TestColumn1", typeof(int)));
        dataTable.Columns.Add(new DataColumn("TestColumn2", typeof(decimal)));
        dataTable.Columns.Add(new DataColumn("TestColumn3", typeof(string)));
        dataTable.Columns.Add(new DataColumn("TestColumn4", typeof(bool)));
        dataTable.Rows.Add(3, 20.7m, "abc", false);
        dataTable.Rows.Add(1, 10.5m, "jkl", true);
        dataTable.Rows.Add(null, null, null, null);
        dataTable.Rows.Add(2, 30.9m, "def", false);

        var totalPanel = new ExcelTotalsPanel(dataTable, Substitute.For<IXLNamedRange>(), Substitute.For<object>(),
            new TestReport().TemplateProcessor);
        IEnumerator enumerator = EnumeratorFactory.Create(dataTable);
        IList<ExcelTotalsPanel.ParsedAggregationFunc> totalCells = new List<ExcelTotalsPanel.ParsedAggregationFunc>
        {
            new(AggregateFunction.Sum, "di:TestColumn1"),
            new(AggregateFunction.Sum, "di:TestColumn2"),
            new(AggregateFunction.Sum, "di:TestColumn3"),
            new(AggregateFunction.Count, "di:TestColumn1"),
            new(AggregateFunction.Count, "di:TestColumn3"),
            new(AggregateFunction.Avg, "di:TestColumn1"),
            new(AggregateFunction.Avg, "di:TestColumn2"),
            new(AggregateFunction.Min, "di:TestColumn1"),
            new(AggregateFunction.Max, "di:TestColumn1"),
            new(AggregateFunction.Min, "di:TestColumn2"),
            new(AggregateFunction.Max, "di:TestColumn2"),
            new(AggregateFunction.Min, "di:TestColumn3"),
            new(AggregateFunction.Max, "di:TestColumn3"),
            new(AggregateFunction.Min, "di:TestColumn4"),
            new(AggregateFunction.Max, "di:TestColumn4")
        };

        var method = totalPanel.GetType().GetMethod("DoAggregation", BindingFlags.Instance | BindingFlags.NonPublic);
        method.Invoke(totalPanel, new object[] {enumerator, totalCells, null});

        Assert.AreEqual(6, totalCells[0].Result);
        Assert.AreEqual(62.1m, totalCells[1].Result);
        Assert.AreEqual("abcjkldef", totalCells[2].Result);
        Assert.AreEqual(4, totalCells[3].Result);
        Assert.AreEqual(4, totalCells[4].Result);
        Assert.AreEqual((double) 6 / 4, totalCells[5].Result);
        Assert.AreEqual(62.1 / 4, totalCells[6].Result);
        Assert.AreEqual(1, totalCells[7].Result);
        Assert.AreEqual(3, totalCells[8].Result);
        Assert.AreEqual(10.5m, totalCells[9].Result);
        Assert.AreEqual(30.9m, totalCells[10].Result);
        Assert.AreEqual("abc", totalCells[11].Result);
        Assert.AreEqual("jkl", totalCells[12].Result);
        Assert.AreEqual(false, totalCells[13].Result);
        Assert.AreEqual(true, totalCells[14].Result);

        // Reset all results before next test
        foreach (var totalCell in totalCells)
        {
            totalCell.Result = null;
        }

        var data = GetTestData();

        enumerator = EnumeratorFactory.Create(data);
        totalCells.Add(new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Sum, "di:Result.Amount"));

        method.Invoke(totalPanel, new object[] {enumerator, totalCells, null});

        Assert.AreEqual(6, totalCells[0].Result);
        Assert.AreEqual(62.1m, totalCells[1].Result);
        Assert.AreEqual("abcjkldef", totalCells[2].Result);
        Assert.AreEqual(4, totalCells[3].Result);
        Assert.AreEqual(4, totalCells[4].Result);
        Assert.AreEqual((double) 6 / 4, totalCells[5].Result);
        Assert.AreEqual(62.1 / 4, totalCells[6].Result);
        Assert.AreEqual(1, totalCells[7].Result);
        Assert.AreEqual(3, totalCells[8].Result);
        Assert.AreEqual(10.5m, totalCells[9].Result);
        Assert.AreEqual(30.9m, totalCells[10].Result);
        Assert.AreEqual("abc", totalCells[11].Result);
        Assert.AreEqual("jkl", totalCells[12].Result);
        Assert.AreEqual(false, totalCells[13].Result);
        Assert.AreEqual(true, totalCells[14].Result);
        Assert.AreEqual(410.59m, totalCells[15].Result);
    }

    [Test]
    public void TestDoAggregationWithEmptyData()
    {
        var data = new List<Test>();
        var totalPanel = new ExcelTotalsPanel(data, Substitute.For<IXLNamedRange>(), Substitute.For<object>(),
            Substitute.For<ITemplateProcessor>());
        IEnumerator enumerator = EnumeratorFactory.Create(data);
        IList<ExcelTotalsPanel.ParsedAggregationFunc> totalCells = new List<ExcelTotalsPanel.ParsedAggregationFunc>
        {
            new(AggregateFunction.Sum, "TestColumn1"),
            new(AggregateFunction.Count, "TestColumn1"),
            new(AggregateFunction.Avg, "TestColumn1"),
            new(AggregateFunction.Min, "TestColumn1"),
            new(AggregateFunction.Max, "TestColumn1")
        };

        var method = totalPanel.GetType().GetMethod("DoAggregation", BindingFlags.Instance | BindingFlags.NonPublic);
        method.Invoke(totalPanel, new object[] {enumerator, totalCells, null});

        Assert.AreEqual(0, totalCells[0].Result);
        Assert.AreEqual(0, totalCells[1].Result);
        Assert.AreEqual(0, totalCells[2].Result);
        Assert.IsNull(totalCells[3].Result);
        Assert.IsNull(totalCells[4].Result);
    }

    [Test]
    public void TestDoAggregationWithIntData()
    {
        var data = new DataProvider().GetIntData();

        var totalPanel = new ExcelTotalsPanel(data, Substitute.For<IXLNamedRange>(), Substitute.For<object>(),
            new TestReport().TemplateProcessor);
        IEnumerator enumerator = EnumeratorFactory.Create(data);
        IList<ExcelTotalsPanel.ParsedAggregationFunc> totalCells = new List<ExcelTotalsPanel.ParsedAggregationFunc>
        {
            new(AggregateFunction.Sum, "di:di"),
            new(AggregateFunction.Avg, "di:di")
        };

        var method = totalPanel.GetType().GetMethod("DoAggregation", BindingFlags.Instance | BindingFlags.NonPublic);
        method.Invoke(totalPanel, new object[] {enumerator, totalCells, null});

        Assert.AreEqual(55, totalCells[0].Result);
        Assert.AreEqual(5.5, totalCells[1].Result);
    }

    [Test]
    public void TestDoAggregationWithBadData()
    {
        var data = GetTestData();

        var totalPanel = new ExcelTotalsPanel(data, Substitute.For<IXLNamedRange>(), Substitute.For<object>(),
            new TestReport().TemplateProcessor);
        IEnumerator enumerator = EnumeratorFactory.Create(data);
        IList<ExcelTotalsPanel.ParsedAggregationFunc> totalCells = new List<ExcelTotalsPanel.ParsedAggregationFunc>
        {
            new(AggregateFunction.Sum, "di:TestColumn4")
        };

        var method = totalPanel.GetType().GetMethod("DoAggregation", BindingFlags.Instance | BindingFlags.NonPublic);
        ExceptionAssert.ThrowsBaseException<RuntimeBinderException>(() =>
            method.Invoke(totalPanel, new object[] {enumerator, totalCells, null}));

        enumerator.Reset();
        totalCells = new List<ExcelTotalsPanel.ParsedAggregationFunc>
        {
            new(AggregateFunction.Min, "di:BadColumn")
        };
        ExceptionAssert.ThrowsBaseException<InvalidOperationException>(
            () => method.Invoke(totalPanel, new object[] {enumerator, totalCells, null}),
            "For Min and Max aggregation functions data items must implement IComparable interface");

        enumerator.Reset();
        totalCells = new List<ExcelTotalsPanel.ParsedAggregationFunc>
        {
            new((AggregateFunction) 6, "di:TestColumn1")
        };
        ExceptionAssert.ThrowsBaseException<NotSupportedException>(
            () => method.Invoke(totalPanel, new object[] {enumerator, totalCells, null}),
            "Unsupportable aggregation function");
    }

    [Test]
    public void TestCustomAggregation()
    {
        var data = GetTestData();

        var totalPanel = new ExcelTotalsPanel(data, Substitute.For<IXLNamedRange>(), new TestReportForAggregation(),
            new TestReport().TemplateProcessor);
        IEnumerator enumerator = EnumeratorFactory.Create(data);
        IList<ExcelTotalsPanel.ParsedAggregationFunc> totalCells = new List<ExcelTotalsPanel.ParsedAggregationFunc>
        {
            new(AggregateFunction.Custom, "di:TestColumn2")
            {
                CustomFunc = "CustomAggregation"
            }
        };

        var method = totalPanel.GetType().GetMethod("DoAggregation", BindingFlags.Instance | BindingFlags.NonPublic);
        method.Invoke(totalPanel, new object[] {enumerator, totalCells, null});
        Assert.AreEqual(24.18125m, totalCells.First().Result);

        enumerator.Reset();
        totalCells.First().CustomFunc = null;
        ExceptionAssert.ThrowsBaseException<InvalidOperationException>(
            () => method.Invoke(totalPanel, new object[] {enumerator, totalCells, null}),
            "The custom type of aggregation is specified in the template but custom function is missing");

        enumerator.Reset();
        totalCells.First().CustomFunc = string.Empty;
        ExceptionAssert.ThrowsBaseException<InvalidOperationException>(
            () => method.Invoke(totalPanel, new object[] {enumerator, totalCells, null}),
            "The custom type of aggregation is specified in the template but custom function is missing");

        enumerator.Reset();
        totalCells.First().CustomFunc = " ";
        ExceptionAssert.ThrowsBaseException<InvalidOperationException>(
            () => method.Invoke(totalPanel, new object[] {enumerator, totalCells, null}),
            "The custom type of aggregation is specified in the template but custom function is missing");

        enumerator.Reset();
        totalCells.First().CustomFunc = "BadMethod";
        ExceptionAssert.ThrowsBaseException<MethodNotFoundException>(
            () => method.Invoke(totalPanel, new object[] {enumerator, totalCells, null}),
            $"Cannot find public instance method \"BadMethod\" in type \"{nameof(TestReportForAggregation)}\"");
    }

    [Test]
    public void TestAggregationPostOperation()
    {
        var data = GetTestData();

        var totalPanel = new ExcelTotalsPanel(data, Substitute.For<IXLNamedRange>(), new TestReportForAggregation(),
            new TestReport().TemplateProcessor);
        IEnumerator enumerator = EnumeratorFactory.Create(data);
        IList<ExcelTotalsPanel.ParsedAggregationFunc> totalCells = new List<ExcelTotalsPanel.ParsedAggregationFunc>
        {
            new(AggregateFunction.Sum, "di:TestColumn2") {PostProcessFunction = "PostSumOperation"},
            new(AggregateFunction.Min, "di:TestColumn3") {PostProcessFunction = "PostMinOperation"},
            new(AggregateFunction.Custom, "di:TestColumn2")
            {
                CustomFunc = "CustomAggregation",
                PostProcessFunction = "PostCustomAggregation"
            }
        };

        var method = totalPanel.GetType().GetMethod("DoAggregation", BindingFlags.Instance | BindingFlags.NonPublic);
        method.Invoke(totalPanel, new object[] {enumerator, totalCells, null});
        Assert.AreEqual(22.033.ToString("F3"), totalCells[0].Result);
        Assert.AreEqual("ABC", totalCells[1].Result);
        Assert.AreEqual(24, totalCells[2].Result);
    }

    [Test]
    public void TestParseTotalCells()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");

        var range = ws.Range(1, 1, 1, 7);
        range.AddToNamed("Test", XLScope.Worksheet);

        ws.Cell(1, 1).Value = "Plain text";
        ws.Cell(1, 2).Value = "{Sum(di:Amount)}";
        ws.Cell(1, 3).Value = "{ Custom(DI:Amount, CustomFunc)  }";
        ws.Cell(1, 4).Value = "{Min(di:Value, CustomFunc, PostFunc)}";
        ws.Cell(1, 5).Value =
            "Text {count(di:Number)} {p:Text} {AVG( di:Value, ,  PostFunc )} {Text} Text {Max(di:Val)} {Min(val)}";
        ws.Cell(1, 6).Value = "{Mix(di:Amount)}";
        ws.Cell(1, 7).Value =
            "Text {Plain Text} Sum(di:Count) {sf:Format(Sum(di:Amount,,PostAggregation), #,,0.00)} {p:Text} {Max(di:Count)} {m:Meth(1, Avg( di : Value ), Min(di:Amount, CustomAggregation, PostAggregation), \"Str\"} {sv:RenderDate} m:Meth2(Avg(di:Value))";
        ws.Cell(1, 8).Value = "{Sum(di:Amount)}";

        var panel = new ExcelTotalsPanel("Stub", ws.NamedRange("Test"), report, report.TemplateProcessor);
        var method = panel.GetType().GetMethod("ParseTotalCells", BindingFlags.Instance | BindingFlags.NonPublic);
        var result = (IDictionary<IXLCell, IList<ExcelTotalsPanel.ParsedAggregationFunc>>) method.Invoke(panel, null);

        Assert.AreEqual(5, result.Count);
        Assert.AreEqual("Plain text", ws.Cell(1, 1).Value);
        Assert.AreEqual("{Mix(di:Amount)}", ws.Cell(1, 6).Value);
        Assert.AreEqual("{Sum(di:Amount)}", ws.Cell(1, 8).Value);

        Assert.IsTrue(Regex.IsMatch(ws.Cell(1, 2).Value.ToString(), @"{di:AggFunc_[0-9a-f]{32}}"));
        Assert.AreEqual(1, result[ws.Cell(1, 2)].Count);
        Assert.AreEqual(AggregateFunction.Sum, result[ws.Cell(1, 2)].First().AggregateFunction);
        Assert.AreEqual("di:Amount", result[ws.Cell(1, 2)].First().ColumnName);
        Assert.IsNull(result[ws.Cell(1, 2)].First().CustomFunc);
        Assert.IsNull(result[ws.Cell(1, 2)].First().PostProcessFunction);
        Assert.IsNull(result[ws.Cell(1, 2)].First().Result);
        Assert.IsTrue(Guid.TryParse(result[ws.Cell(1, 2)].First().UniqueName, out _));

        Assert.IsTrue(Regex.IsMatch(ws.Cell(1, 3).Value.ToString(), @"{ di:AggFunc_[0-9a-f]{32}  }"));
        Assert.AreEqual(1, result[ws.Cell(1, 3)].Count);
        Assert.AreEqual(AggregateFunction.Custom, result[ws.Cell(1, 3)].First().AggregateFunction);
        Assert.AreEqual("DI:Amount", result[ws.Cell(1, 3)].First().ColumnName);
        Assert.AreEqual("CustomFunc", result[ws.Cell(1, 3)].First().CustomFunc);
        Assert.IsNull(result[ws.Cell(1, 3)].First().PostProcessFunction);
        Assert.IsNull(result[ws.Cell(1, 3)].First().Result);
        Assert.IsTrue(Guid.TryParse(result[ws.Cell(1, 3)].First().UniqueName, out _));

        Assert.IsTrue(Regex.IsMatch(ws.Cell(1, 4).Value.ToString(), @"{di:AggFunc_[0-9a-f]{32}}"));
        Assert.AreEqual(1, result[ws.Cell(1, 4)].Count);
        Assert.AreEqual(AggregateFunction.Min, result[ws.Cell(1, 4)].First().AggregateFunction);
        Assert.AreEqual("di:Value", result[ws.Cell(1, 4)].First().ColumnName);
        Assert.AreEqual("CustomFunc", result[ws.Cell(1, 4)].First().CustomFunc);
        Assert.AreEqual("PostFunc", result[ws.Cell(1, 4)].First().PostProcessFunction);
        Assert.IsNull(result[ws.Cell(1, 4)].First().Result);
        Assert.IsTrue(Guid.TryParse(result[ws.Cell(1, 4)].First().UniqueName, out _));

        Assert.IsTrue(Regex.IsMatch(ws.Cell(1, 5).Value.ToString(),
            @"Text {di:AggFunc_[0-9a-f]{32}} {p:Text} {di:AggFunc_[0-9a-f]{32}} {Text} Text {di:AggFunc_[0-9a-f]{32}} {Min\(val\)}"));
        Assert.AreEqual(3, result[ws.Cell(1, 5)].Count);
        Assert.AreEqual(AggregateFunction.Count, result[ws.Cell(1, 5)][0].AggregateFunction);
        Assert.AreEqual("di:Number", result[ws.Cell(1, 5)][0].ColumnName);
        Assert.IsNull(result[ws.Cell(1, 5)][0].CustomFunc);
        Assert.IsNull(result[ws.Cell(1, 5)][0].PostProcessFunction);
        Assert.IsNull(result[ws.Cell(1, 5)][0].Result);
        Assert.IsTrue(Guid.TryParse(result[ws.Cell(1, 5)][0].UniqueName, out _));
        Assert.AreEqual(AggregateFunction.Avg, result[ws.Cell(1, 5)][1].AggregateFunction);
        Assert.AreEqual("di:Value", result[ws.Cell(1, 5)][1].ColumnName);
        Assert.IsNull(result[ws.Cell(1, 5)][1].CustomFunc);
        Assert.AreEqual("PostFunc", result[ws.Cell(1, 5)][1].PostProcessFunction);
        Assert.IsNull(result[ws.Cell(1, 5)][1].Result);
        Assert.IsTrue(Guid.TryParse(result[ws.Cell(1, 5)][1].UniqueName, out _));
        Assert.AreEqual(AggregateFunction.Max, result[ws.Cell(1, 5)][2].AggregateFunction);
        Assert.AreEqual("di:Val", result[ws.Cell(1, 5)][2].ColumnName);
        Assert.IsNull(result[ws.Cell(1, 5)][2].CustomFunc);
        Assert.IsNull(result[ws.Cell(1, 5)][2].PostProcessFunction);
        Assert.IsNull(result[ws.Cell(1, 5)][2].Result);
        Assert.IsTrue(Guid.TryParse(result[ws.Cell(1, 5)][2].UniqueName, out _));

        Assert.IsTrue(Regex.IsMatch(ws.Cell(1, 7).Value.ToString(),
            @"Text {Plain Text} Sum\(di:Count\) {sf:Format\(di:AggFunc_[0-9a-f]{32}, #,,0.00\)} {p:Text} {di:AggFunc_[0-9a-f]{32}} {m:Meth\(1, di:AggFunc_[0-9a-f]{32}, di:AggFunc_[0-9a-f]{32}, ""Str""} {sv:RenderDate} m:Meth2\(Avg\(di:Value\)\)"));
        Assert.AreEqual(4, result[ws.Cell(1, 7)].Count);
        Assert.AreEqual(AggregateFunction.Sum, result[ws.Cell(1, 7)][0].AggregateFunction);
        Assert.AreEqual("di:Amount", result[ws.Cell(1, 7)][0].ColumnName);
        Assert.IsNull(result[ws.Cell(1, 7)][0].CustomFunc);
        Assert.AreEqual("PostAggregation", result[ws.Cell(1, 7)][0].PostProcessFunction);
        Assert.IsNull(result[ws.Cell(1, 7)][0].Result);
        Assert.IsTrue(Guid.TryParse(result[ws.Cell(1, 7)][0].UniqueName, out _));
        Assert.AreEqual(AggregateFunction.Max, result[ws.Cell(1, 7)][1].AggregateFunction);
        Assert.AreEqual("di:Count", result[ws.Cell(1, 7)][1].ColumnName);
        Assert.IsNull(result[ws.Cell(1, 7)][1].CustomFunc);
        Assert.IsNull(result[ws.Cell(1, 7)][1].PostProcessFunction);
        Assert.IsNull(result[ws.Cell(1, 7)][1].Result);
        Assert.IsTrue(Guid.TryParse(result[ws.Cell(1, 7)][1].UniqueName, out _));
        Assert.AreEqual(AggregateFunction.Avg, result[ws.Cell(1, 7)][2].AggregateFunction);
        Assert.AreEqual("di : Value", result[ws.Cell(1, 7)][2].ColumnName);
        Assert.IsNull(result[ws.Cell(1, 7)][2].CustomFunc);
        Assert.IsNull(result[ws.Cell(1, 7)][2].PostProcessFunction);
        Assert.IsNull(result[ws.Cell(1, 7)][2].Result);
        Assert.IsTrue(Guid.TryParse(result[ws.Cell(1, 7)][2].UniqueName, out _));
        Assert.AreEqual(AggregateFunction.Min, result[ws.Cell(1, 7)][3].AggregateFunction);
        Assert.AreEqual("di:Amount", result[ws.Cell(1, 7)][3].ColumnName);
        Assert.AreEqual("CustomAggregation", result[ws.Cell(1, 7)][3].CustomFunc);
        Assert.AreEqual("PostAggregation", result[ws.Cell(1, 7)][3].PostProcessFunction);
        Assert.IsNull(result[ws.Cell(1, 7)][3].Result);
        Assert.IsTrue(Guid.TryParse(result[ws.Cell(1, 7)][3].UniqueName, out _));
    }

    [Test]
    public void TestParseTotalCellsErrors()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");

        var range = ws.Range(1, 1, 1, 1);
        range.AddToNamed("Test", XLScope.Worksheet);

        var templateProcessor = Substitute.For<ITemplateProcessor>();
        templateProcessor.LeftTemplateBorder.Returns("<");
        templateProcessor.RightTemplateBorder.Returns(">");
        templateProcessor.MemberLabelSeparator.Returns("-");
        templateProcessor.DataItemMemberLabel.Returns("d");

        var report = new TestReport
        {
            TemplateProcessor = templateProcessor,
            Workbook = wb
        };

        var panel = new ExcelTotalsPanel("Stub", ws.NamedRange("Test"), report, report.TemplateProcessor);
        var method = panel.GetType().GetMethod("ParseTotalCells", BindingFlags.Instance | BindingFlags.NonPublic);

        ws.Cell(1, 1).Value = "<Sum(d-Val, fn1, fn2, fn3)>";
        ExceptionAssert.ThrowsBaseException<InvalidOperationException>(() => method.Invoke(panel, null),
            "Aggregation function must have at least one but no more than 3 parameters");
    }

    private IList<Test> GetTestData()
    {
        return new List<Test>
        {
            new(3, 20.7m, "abc", false) {Result = new ComplexType {Amount = 155.05m}},
            new(1, 10.5m, "jkl", true) {Result = new ComplexType()},
            new(null, null, null, null) {Result = new ComplexType()},
            new(2, 30.9m, "def", false) {Result = new ComplexType {Amount = 255.54m}}
        };
    }

    private class Test
    {
        public Test(int? testColumn1, decimal? testColumn2, string testColumn3, bool? testColumn4)
        {
            TestColumn1 = testColumn1;
            TestColumn2 = testColumn2;
            TestColumn3 = testColumn3;
            TestColumn4 = testColumn4;
            BadColumn = new Test();
        }

        private Test()
        {
        }

        public int? TestColumn1 { get; }
        public decimal? TestColumn2 { get; }
        public string TestColumn3 { get; }
        public bool? TestColumn4 { get; }
        public Test BadColumn { get; }
        public ComplexType Result { get; set; }
    }

    private class ComplexType
    {
        public decimal Amount { get; set; }
    }

    private class TestReportForAggregation
    {
        public decimal CustomAggregation(decimal result, decimal currentValue, int itemNumber)
        {
            return (result + currentValue) / 2 + itemNumber;
        }

        public string PostSumOperation(decimal result, int itemsCount)
        {
            return ((result + itemsCount) / 3).ToString("F3");
        }

        public string PostMinOperation(string result, int itemsCount)
        {
            return result.ToUpper();
        }

        public int PostCustomAggregation(decimal result, int itemsCount)
        {
            return (int) decimal.Round(result, 0);
        }
    }
}