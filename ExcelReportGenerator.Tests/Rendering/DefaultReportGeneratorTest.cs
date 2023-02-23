using System.Reflection;
using ClosedXML.Excel;
using ExcelReportGenerator.Rendering;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Rendering.Providers;
using ExcelReportGenerator.Rendering.Providers.DataItemValueProviders;
using ExcelReportGenerator.Rendering.Providers.VariableProviders;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using ExcelReportGenerator.Tests.CustomAsserts;
using ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests;

namespace ExcelReportGenerator.Tests.Rendering;

public class DefaultReportGeneratorTest
{
    [Test]
    public void TestMakePanelsHierarchy()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");

        var panel1Range = ws.Range(1, 1, 4, 4);
        var panel2Range = ws.Range(1, 1, 2, 4);
        panel2Range.AddToNamed("Panel2", XLScope.Worksheet);
        var panel3Range = ws.Range(2, 1, 2, 4);
        panel3Range.AddToNamed("Panel3", XLScope.Workbook);
        var panel4Range = ws.Range(5, 1, 6, 5);
        var panel5Range = ws.Range(6, 1, 6, 5);
        panel5Range.AddToNamed("Panel5", XLScope.Worksheet);
        var panel6Range = ws.Range(3, 1, 4, 4);
        var panel7Range = ws.Range(10, 10, 10, 10);
        var panel8Range = ws.Range(8, 9, 9, 10);
        panel8Range.AddToNamed("Panel8", XLScope.Worksheet);

        var panel1 = new ExcelPanel(panel1Range, new object(), Substitute.For<ITemplateProcessor>());
        var panel2 = new ExcelDataSourcePanel("Stub", ws.NamedRange("Panel2"), new object(),
            Substitute.For<ITemplateProcessor>());
        var panel3 = new ExcelDataSourcePanel("Stub", wb.NamedRange("Panel3"), new object(),
            Substitute.For<ITemplateProcessor>());
        var panel4 = new ExcelPanel(panel4Range, new object(), Substitute.For<ITemplateProcessor>());
        var panel5 = new ExcelDataSourceDynamicPanel("Stub", ws.NamedRange("Panel5"), new object(),
            Substitute.For<ITemplateProcessor>());
        var panel6 = new ExcelPanel(panel6Range, new object(), Substitute.For<ITemplateProcessor>());
        var panel7 = new ExcelPanel(panel7Range, new object(), Substitute.For<ITemplateProcessor>());
        var panel8 = new ExcelTotalsPanel("Stub", ws.NamedRange("Panel8"), new object(),
            Substitute.For<ITemplateProcessor>());

        IDictionary<string, (IExcelPanel, string)> panelsFlatView = new Dictionary<string, (IExcelPanel, string)>
        {
            ["Panel1"] = (panel1, null),
            ["Panel2"] = (panel2, "Panel1"),
            ["Panel3"] = (panel3, "Panel2"),
            ["Panel4"] = (panel4, null),
            ["Panel5"] = (panel5, "Panel4"),
            ["Panel6"] = (panel6, "Panel1"),
            ["Panel7"] = (panel7, null),
            ["Panel8"] = (panel8, null)
        };

        var reportGenerator = new DefaultReportGenerator(new object());
        var method = reportGenerator.GetType()
            .GetMethod("MakePanelsHierarchy", BindingFlags.Instance | BindingFlags.NonPublic);

        var rootPanel = new ExcelPanel(ws.Range(1, 1, 10, 10), new object(), Substitute.For<ITemplateProcessor>());
        method.Invoke(reportGenerator, new object[] {panelsFlatView, rootPanel});

        Assert.AreEqual(4, rootPanel.Children.Count);
        Assert.AreEqual(panel1Range, rootPanel.Children[0].Range);
        Assert.AreEqual(panel4Range, rootPanel.Children[1].Range);
        Assert.AreEqual(panel7Range, rootPanel.Children[2].Range);
        Assert.AreEqual(ws.NamedRange("Panel8").Ranges.First(), rootPanel.Children[3].Range);
        Assert.AreEqual(rootPanel, panel1.Parent);
        Assert.AreEqual(rootPanel, panel4.Parent);
        Assert.AreEqual(rootPanel, panel7.Parent);
        Assert.AreEqual(rootPanel, panel8.Parent);
        Assert.IsNull(rootPanel.Parent);

        Assert.AreEqual(2, panel1.Children.Count);
        Assert.AreEqual(ws.NamedRange("Panel2").Ranges.First(), panel1.Children[0].Range);
        Assert.AreEqual(panel6Range, panel1.Children[1].Range);
        Assert.AreEqual(panel1, panel2.Parent);
        Assert.AreEqual(panel1, panel6.Parent);

        Assert.AreEqual(1, panel2.Children.Count);
        Assert.AreEqual(wb.NamedRange("Panel3").Ranges.First(), panel2.Children[0].Range);
        Assert.AreEqual(panel2, panel3.Parent);

        Assert.AreEqual(1, panel4.Children.Count);
        Assert.AreEqual(ws.NamedRange("Panel5").Ranges.First(), panel4.Children[0].Range);
        Assert.AreEqual(panel4, panel5.Parent);
    }

    [Test]
    public void TestMakePanelsHierarchyWithBadParent()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");

        var panel1Range = ws.Range(1, 1, 4, 4);
        var panel2Range = ws.Range(1, 1, 2, 4);
        panel2Range.AddToNamed("Panel2", XLScope.Worksheet);

        var panel1 = new ExcelPanel(panel1Range, new object(), Substitute.For<ITemplateProcessor>());
        var panel2 = new ExcelDataSourcePanel("Stub", ws.NamedRange("Panel2"), new object(),
            Substitute.For<ITemplateProcessor>());

        IDictionary<string, (IExcelPanel, string)> panelsFlatView = new Dictionary<string, (IExcelPanel, string)>
        {
            ["Panel1"] = (panel1, null),
            ["Panel2"] = (panel2, "panel1")
        };

        var reportGenerator = new DefaultReportGenerator(new object());
        var method = reportGenerator.GetType()
            .GetMethod("MakePanelsHierarchy", BindingFlags.Instance | BindingFlags.NonPublic);

        var rootPanel = new ExcelPanel(ws.Range(1, 1, 10, 10), new object(), Substitute.For<ITemplateProcessor>());
        ExceptionAssert.ThrowsBaseException<InvalidOperationException>(
            () => method.Invoke(reportGenerator, new object[] {panelsFlatView, rootPanel}),
            "Cannot find parent panel with name \"panel1\" for panel \"Panel2\"");

        var panel3Range = ws.Range(3, 1, 5, 4);
        panel3Range.AddToNamed("Panel3", XLScope.Worksheet);

        var panel3 = new ExcelDataSourcePanel("Stub", ws.NamedRange("Panel3"), new object(),
            Substitute.For<ITemplateProcessor>());

        panelsFlatView = new Dictionary<string, (IExcelPanel, string)>
        {
            ["Panel1"] = (panel1, null),
            ["Panel3"] = (panel3, "Panel1")
        };

        ExceptionAssert.ThrowsBaseException<InvalidOperationException>(
            () => method.Invoke(reportGenerator, new object[] {panelsFlatView, rootPanel}),
            $"Panel \"{panel1Range}\" is not a parent of the panel \"Panel3\". Child range is outside of the parent range.");
    }

    [Test]
    public void TestGetPanelsNamedRanges()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");

        var panel1Range = ws.Range(1, 1, 4, 4);
        panel1Range.AddToNamed("s_Panel1", XLScope.Worksheet);
        var panel2Range = ws.Range(1, 1, 2, 4);
        panel2Range.AddToNamed("D_Panel2", XLScope.Worksheet);
        var panel3Range = ws.Range(2, 1, 2, 4);
        panel3Range.AddToNamed("DYN_Panel3", XLScope.Worksheet);
        var panel4Range = ws.Range(5, 1, 6, 5);
        panel4Range.AddToNamed("t_Panel4", XLScope.Worksheet);
        var panel5Range = ws.Range(6, 1, 6, 5);
        panel5Range.AddToNamed("ss_Panel5", XLScope.Worksheet);
        var panel6Range = ws.Range(3, 1, 4, 4);
        panel6Range.AddToNamed("S_Panel6", XLScope.Worksheet);
        var panel7Range = ws.Range(10, 10, 10, 10);
        panel7Range.AddToNamed("d-Panel7", XLScope.Worksheet);
        var panel8Range = ws.Range(8, 9, 9, 10);
        panel8Range.AddToNamed("d_Panel8", XLScope.Worksheet);
        var panel9Range = ws.Range(11, 11, 11, 11);
        panel9Range.AddToNamed("_d_Panel9 ", XLScope.Worksheet);

        var reportGenerator = new DefaultReportGenerator(new object());
        var method = reportGenerator.GetType()
            .GetMethod("GetPanelsNamedRanges", BindingFlags.Instance | BindingFlags.NonPublic);

        var result = (IList<IXLNamedRange>) method.Invoke(reportGenerator, new object[] {ws.NamedRanges});

        Assert.AreEqual(6, result.Count);
        Assert.AreEqual("s_Panel1", result[0].Name);
        Assert.AreEqual("D_Panel2", result[1].Name);
        Assert.AreEqual("DYN_Panel3", result[2].Name);
        Assert.AreEqual("t_Panel4", result[3].Name);
        Assert.AreEqual("S_Panel6", result[4].Name);
        Assert.AreEqual("d_Panel8", result[5].Name);
    }

    [Test]
    public void TestGetRootRange()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");

        var reportGenerator = new DefaultReportGenerator(new object());
        var method = reportGenerator.GetType()
            .GetMethod("GetRootRange", BindingFlags.Instance | BindingFlags.NonPublic);

        ws.Cell(10, 10).Value = "Val";
        ws.Cell(20, 20).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

        Assert.AreEqual(ws.Range(ws.FirstCell(), ws.Cell(10, 10)),
            method.Invoke(reportGenerator, new object[] {ws, null}));

        var namedRanges = new List<IXLNamedRange>();
        Assert.AreEqual(ws.Range(ws.FirstCell(), ws.Cell(10, 10)),
            method.Invoke(reportGenerator, new object[] {ws, namedRanges}));

        var range = ws.Range(ws.Cell(2, 2), ws.Cell(3, 3));
        range.AddToNamed("range1", XLScope.Worksheet);
        namedRanges.Add(ws.NamedRange("range1"));
        Assert.AreEqual(ws.Range(ws.FirstCell(), ws.Cell(10, 10)),
            method.Invoke(reportGenerator, new object[] {ws, namedRanges}));

        range = ws.Range(ws.Cell(2, 2), ws.Cell(10, 11));
        range.AddToNamed("range2", XLScope.Worksheet);
        namedRanges.Add(ws.NamedRange("range2"));
        Assert.AreEqual(ws.Range(ws.FirstCell(), ws.Cell(10, 11)),
            method.Invoke(reportGenerator, new object[] {ws, namedRanges}));

        range = ws.Range(ws.Cell(5, 5), ws.Cell(11, 10));
        range.AddToNamed("range3", XLScope.Worksheet);
        namedRanges.Add(ws.NamedRange("range3"));
        Assert.AreEqual(ws.Range(ws.FirstCell(), ws.Cell(11, 11)),
            method.Invoke(reportGenerator, new object[] {ws, namedRanges}));

        range = ws.Range(ws.Cell(18, 18), ws.Cell(18, 18));
        range.AddToNamed("range4", XLScope.Worksheet);
        namedRanges.Add(ws.NamedRange("range4"));
        Assert.AreEqual(ws.Range(ws.FirstCell(), ws.Cell(18, 18)),
            method.Invoke(reportGenerator, new object[] {ws, namedRanges}));

        ws.Cell(21, 21).Value = "Val2";
        Assert.AreEqual(ws.Range(ws.FirstCell(), ws.Cell(21, 21)),
            method.Invoke(reportGenerator, new object[] {ws, namedRanges}));
    }

    [Test]
    public void TestRender()
    {
        var report = new TestReport();
        var wb = report.Workbook;
        var sheet1 = wb.AddWorksheet("Sheet1");
        var sheet2 = wb.AddWorksheet("Sheet2");
        var reportGenerator = new TestReportGenerator(report);

        var parentRange = sheet1.Range(2, 2, 3, 5);
        parentRange.AddToNamed("d_Parent", XLScope.Worksheet);
        var parentNamedRange = sheet1.NamedRange("d_Parent");
        parentNamedRange.Comment = "DataSource = m:DataProvider:GetIEnumerable();";

        var childRange = sheet1.Range(3, 2, 3, 5);
        childRange.AddToNamed("d_Child", XLScope.Workbook);
        var childNamedRange = wb.NamedRange("d_Child");
        childNamedRange.Comment =
            $"ParentPanel = d_Parent{Environment.NewLine}DataSource = m:DataProvider:GetChildIEnumerable(di:Name);GroupBy=1,4";

        sheet1.Cell(2, 2).Value = "{di:Name}";
        sheet1.Cell(2, 3).Value = "{di:Date}";
        sheet1.Cell(3, 3).Value = "{di:Field1}";
        sheet1.Cell(3, 4).Value = "{di:Field2}";
        sheet1.Cell(3, 5).Value = "{di:parent:Sum}";

        sheet1.Cell(3, 5).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        sheet1.Cell(3, 5).Style.Font.Bold = true;

        var simpleRange = sheet1.Range(2, 7, 3, 8);
        simpleRange.AddToNamed("s_Simple", XLScope.Worksheet);

        sheet1.Cell(2, 7).Value = "{p:StrParam}";
        sheet1.Cell(3, 8).Value = "{p:IntParam}";

        sheet1.Cell(2, 10).Value = "{sv:SheetName}";
        sheet1.Cell(3, 10).Value = "{sv:SheetNumber}";

        var dynamicRange = sheet2.Range(2, 2, 4, 2);
        dynamicRange.AddToNamed("dyn_Dynamic", XLScope.Workbook);
        var dynamicNamedRange = wb.NamedRange("dyn_Dynamic");
        dynamicNamedRange.Comment = "DataSource = m:DataProvider:GetAllCustomersDataReader()";

        sheet2.Cell(2, 2).Value = "{Headers}";
        sheet2.Cell(3, 2).Value = "{Data}";
        sheet2.Cell(4, 2).Value = "{Totals}";

        var totalsRange = sheet2.Range(6, 2, 6, 7);
        totalsRange.AddToNamed("t_Totals", XLScope.Worksheet);
        var totalsNamedRange = sheet2.NamedRange("t_Totals");
        totalsNamedRange.Comment =
            "DataSource = m:DataProvider:GetIEnumerable(); BeforeRenderMethodName = TestExcelTotalsPanelBeforeRender; AfterRenderMethodName = TestExcelTotalsPanelAfterRender";

        var intRange = sheet2.Range(10, 2, 10, 2);
        intRange.AddToNamed("d_IntData", XLScope.Worksheet);
        var intNamedRange = sheet2.NamedRange("d_IntData");
        intNamedRange.Comment = "DataSource = m:DataProvider:GetIntData()";

        sheet2.Cell(10, 2).Value = "{di:di}";

        sheet2.Cell(6, 2).Value = "Plain text";
        sheet2.Cell(6, 3).Value = "{Sum(di:Sum)}";
        sheet2.Cell(6, 4).Value = "{ Custom(DI:Sum, CustomAggregation, PostAggregation)  }";
        sheet2.Cell(6, 5).Value = "{Min(di:Sum)}";
        sheet2.Cell(6, 6).Value =
            "Text1 {count(di:Name)} Text2 {avg(di:Sum, , PostAggregationRound)} Text3 {Max(di:Sum)}";
        sheet2.Cell(6, 7).Value = "{Mix(di:Sum)}";

        sheet2.Cell(10, 10).Value = "Plain text";
        sheet2.Cell(1, 1).Value = " { m:Format ( p:DateParam ) } ";
        sheet2.Cell(7, 1).Value = "{P:BoolParam}";

        sheet2.Cell(2, 10).Value = "{sv:SheetName}";
        sheet2.Cell(3, 10).Value = "{sv:SheetNumber}";

        reportGenerator.Render(wb);

        ExcelAssert.AreWorkbooksContentEquals(
            TestHelper.GetExpectedWorkbook(nameof(DefaultReportGeneratorTest), nameof(TestRender)), wb);

        // wb.SaveAs("test.xlsx");
    }

    [Test]
    public void TestRenderWithEvents()
    {
        var report = new TestReport();
        var wb = report.Workbook;
        var sheet1 = wb.AddWorksheet("Sheet1");
        var reportGenerator = new TestReportGenerator(report);

        reportGenerator.BeforeReportRender += (sender, args) => args.Workbook.AddWorksheet("DynamicSheet");
        reportGenerator.BeforeWorksheetRender += (sender, args) =>
        {
            if (args.Worksheet.Name == "DynamicSheet")
            {
                args.Worksheet.Cell(2, 2).Value = "{sv:SheetNumber}";
                args.Worksheet.Cell(2, 3).Value = "{sv:SheetName}";
            }
        };

        reportGenerator.AfterWorksheetRender += (sender, args) =>
        {
            if (args.Worksheet.Name == "Sheet1")
            {
                args.Worksheet.Cell(1, 1).Value = "Sheet1";
            }
        };

        sheet1.Cell(2, 2).Value = "{sv:SheetNumber}";

        reportGenerator.Render(wb);

        ExcelAssert.AreWorkbooksContentEquals(
            TestHelper.GetExpectedWorkbook(nameof(DefaultReportGeneratorTest), nameof(TestRenderWithEvents)), wb);

        //wb.SaveAs("test.xlsx");
    }

    [Test]
    public void TestRenderPartialWorksheets()
    {
        var report = new TestReport();
        var wb = report.Workbook;
        var sheet1 = wb.AddWorksheet("Sheet1");
        var sheet2 = wb.AddWorksheet("Sheet2");
        var reportGenerator = new TestReportGenerator2(report);

        var parentRange = sheet1.Range(2, 2, 3, 5);
        parentRange.AddToNamed("d_Parent", XLScope.Worksheet);
        var parentNamedRange = sheet1.NamedRange("d_Parent");
        parentNamedRange.Comment = "DataSource = m:DataProvider:GetIEnumerable()";

        var childRange = sheet1.Range(3, 2, 3, 5);
        childRange.AddToNamed("d_Child", XLScope.Workbook);
        var childNamedRange = wb.NamedRange("d_Child");
        childNamedRange.Comment =
            $"ParentPanel = d_Parent{Environment.NewLine}DataSource = m:DataProvider:GetChildIEnumerable(di:Name)";

        sheet1.Cell(2, 2).Value = "{di:Name}";
        sheet1.Cell(2, 3).Value = "{di:Date}";
        sheet1.Cell(3, 3).Value = "{di:Field1}";
        sheet1.Cell(3, 4).Value = "{di:Field2}";
        sheet1.Cell(3, 5).Value = "{di:parent:Sum}";

        var simpleRange = sheet1.Range(2, 7, 3, 8);
        simpleRange.AddToNamed("s_Simple", XLScope.Worksheet);

        sheet1.Cell(2, 7).Value = "{p:StrParam}";
        sheet1.Cell(3, 8).Value = "{p:IntParam}";

        var intRange = sheet1.Range(2, 10, 2, 10);
        intRange.AddToNamed("d_IntData", XLScope.Worksheet);
        var intNamedRange = sheet1.NamedRange("d_IntData");
        intNamedRange.Comment = "DataSource = m:DataProvider:GetIntData()";

        sheet1.Cell(2, 10).Value = "{di:self}";

        var dynamicRange = sheet2.Range(2, 2, 4, 2);
        dynamicRange.AddToNamed("dyn_Dynamic", XLScope.Workbook);
        var dynamicNamedRange = wb.NamedRange("dyn_Dynamic");
        dynamicNamedRange.Comment = "DataSource = m:DataProvider:GetAllCustomersDataReader()";

        sheet2.Cell(2, 2).Value = "{Headers}";
        sheet2.Cell(3, 2).Value = "{Data}";
        sheet2.Cell(4, 2).Value = "{Totals}";

        var totalsRange = sheet2.Range(6, 2, 6, 7);
        totalsRange.AddToNamed("t_Totals", XLScope.Worksheet);
        var totalsNamedRange = sheet2.NamedRange("t_Totals");
        totalsNamedRange.Comment =
            "DataSource = m:DataProvider:GetIEnumerable(); BeforeRenderMethodName = TestExcelTotalsPanelBeforeRender; AfterRenderMethodName = TestExcelTotalsPanelAfterRender";

        sheet2.Cell(6, 2).Value = "Plain text";
        sheet2.Cell(6, 3).Value = "{Sum(di:Sum)}";
        sheet2.Cell(6, 4).Value = "{ Custom(DI:Sum, CustomAggregation, PostAggregation)  }";
        sheet2.Cell(6, 5).Value = "{Min(di:Sum)}";
        sheet2.Cell(6, 6).Value =
            "Text1 {count(Name)} Text2 {avg(di:Sum, , PostAggregationRound)} Text3 {Max(Sum)}";
        sheet2.Cell(6, 7).Value = "{Mix(di:Sum)}";

        sheet2.Cell(10, 10).Value = "Plain text";
        sheet2.Cell(1, 1).Value = " { m:Format ( p:DateParam ) } ";
        sheet2.Cell(7, 1).Value = "{P:BoolParam}";

        reportGenerator.Render(wb, new[] {sheet1});

        ExcelAssert.AreWorkbooksContentEquals(
            TestHelper.GetExpectedWorkbook(nameof(DefaultReportGeneratorTest), nameof(TestRenderPartialWorksheets)),
            wb);

        //wb.SaveAs("test.xlsx");
    }

    [Test]
    public void TestRenderWithCustomVariableAndFunctionProviders()
    {
        var report = new TestReport();
        var wb = report.Workbook;
        var sheet1 = wb.AddWorksheet("Sheet1");
        var reportGenerator = new TestReportGenerator(report)
        {
            SystemVariableProvider = new VariableProvider(),
            SystemFunctionsType = typeof(Functions)
        };

        sheet1.Cell(2, 2).Value = "{sv:SheetName}";
        sheet1.Cell(2, 3).Value = "{sv:SheetNumber}";
        sheet1.Cell(2, 4).Value = "{sv:CustomVar}";
        sheet1.Cell(2, 5).Value = "{sf:Format(p:DateParam, yyyyMMdd)}";
        sheet1.Cell(2, 6).Value = "{sf:StaticFunc(p:IntParam)}";
        sheet1.Cell(2, 7).Value = "{sf:InstanceFunc(p:IntParam)}";

        reportGenerator.Render(wb);

        ExcelAssert.AreWorkbooksContentEquals(
            TestHelper.GetExpectedWorkbook(nameof(DefaultReportGeneratorTest),
                nameof(TestRenderWithCustomVariableAndFunctionProviders)), wb);

        //wb.SaveAs("test.xlsx");
    }

    private class TestReportGenerator : DefaultReportGenerator
    {
        private ITypeProvider _typeProvider;

        public TestReportGenerator(object report) : base(report)
        {
        }

        public override ITypeProvider TypeProvider => _typeProvider ??= new DefaultTypeProvider(
            new[] {Assembly.GetExecutingAssembly()},
            Report.GetType());
    }

    private class TestReportGenerator2 : TestReportGenerator
    {
        private ITemplateProcessor _templateProcessor;

        public TestReportGenerator2(object report) : base(report)
        {
        }

        public override ITemplateProcessor TemplateProcessor =>
            _templateProcessor ??= new TemplateProcessor(PropertyValueProvider,
                SystemVariableProvider, MethodCallValueProvider, DataItemValueProvider);
    }

    private class VariableProvider : SystemVariableProvider
    {
        public string CustomVar { get; set; } = "Custom";
    }

    private class Functions : SystemFunctions
    {
        public static string StaticFunc(int param)
        {
            return (param + 5).ToString();
        }

        public string InstanceFunc(int param)
        {
            return (param + 10).ToString();
        }
    }

    private class TemplateProcessor : DefaultTemplateProcessor
    {
        public TemplateProcessor(IPropertyValueProvider propertyValueProvider,
            SystemVariableProvider systemVariableProvider, IMethodCallValueProvider methodCallValueProvider = null,
            IGenericDataItemValueProvider<HierarchicalDataItem> dataItemValueProvider = null) : base(
            propertyValueProvider, systemVariableProvider, methodCallValueProvider, dataItemValueProvider)
        {
        }

        public override string DataItemSelfTemplate => "self";
    }
}