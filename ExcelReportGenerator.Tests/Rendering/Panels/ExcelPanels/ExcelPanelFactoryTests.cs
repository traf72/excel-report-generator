using System.Reflection;
using ClosedXML.Excel;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Rendering;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using ExcelReportGenerator.Tests.CustomAsserts;
using Assert = NUnit.Framework.Legacy.ClassicAssert;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels;

public class ExcelPanelFactoryTests
{
    [Test]
    public void TestCreateSimplePanel()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");
        var range = ws.Range(ws.Cell(1, 1), ws.Cell(2, 2));
        range.AddToNamed("s_Test", XLScope.Worksheet);
        var namedRange = ws.NamedRange("s_Test");

        var report = new object();
        var templateProcessor = Substitute.For<ITemplateProcessor>();
        var parseSettings = new PanelParsingSettings
        {
            SimplePanelPrefix = "s",
            PanelPrefixSeparator = "_"
        };

        var factory = new ExcelPanelFactory(report, templateProcessor, parseSettings);
        var panel = (ExcelPanel) factory.Create(namedRange, new Dictionary<string, string>
        {
            [nameof(ExcelPanel.Type)] = PanelType.Horizontal.ToString(),
            [nameof(ExcelPanel.ShiftType)] = ShiftType.Row.ToString(),
            [nameof(ExcelPanel.RenderPriority)] = "5",
            [nameof(ExcelPanel.BeforeRenderMethodName)] = "BeforeRenderMethodName",
            [nameof(ExcelPanel.AfterRenderMethodName)] = "AfterRenderMethodName"
        });

        Assert.AreEqual(PanelType.Horizontal, panel.Type);
        Assert.AreEqual(ShiftType.Row, panel.ShiftType);
        Assert.AreEqual(5, panel.RenderPriority);
        Assert.AreEqual("BeforeRenderMethodName", panel.BeforeRenderMethodName);
        Assert.AreEqual("AfterRenderMethodName", panel.AfterRenderMethodName);
        Assert.AreEqual(0, panel.Children.Count);
        Assert.IsNull(panel.Parent);
        Assert.AreEqual(namedRange.Ranges.First(), panel.Range);
        Assert.AreSame(report,
            panel.GetType().GetField("_report", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(panel));
        Assert.AreSame(templateProcessor,
            panel.GetType().GetField("_templateProcessor", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(panel));

        namedRange.Delete();
        range.AddToNamed("SS--Test", XLScope.Workbook);
        namedRange = wb.NamedRange("SS--Test");

        parseSettings.SimplePanelPrefix = "ss";
        parseSettings.PanelPrefixSeparator = "--";
        factory = new ExcelPanelFactory(report, templateProcessor, parseSettings);
        panel = (ExcelPanel) factory.Create(namedRange, null);

        Assert.IsInstanceOf<ExcelPanel>(panel);
        Assert.AreEqual(PanelType.Vertical, panel.Type);
        Assert.AreEqual(ShiftType.Cells, panel.ShiftType);
        Assert.AreEqual(0, panel.RenderPriority);
        Assert.AreEqual(0, panel.Children.Count);
        Assert.IsNull(panel.BeforeRenderMethodName);
        Assert.IsNull(panel.AfterRenderMethodName);
        Assert.IsNull(panel.Parent);
        Assert.AreEqual(namedRange.Ranges.First(), panel.Range);
        Assert.AreSame(report,
            panel.GetType().GetField("_report", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(panel));
        Assert.AreSame(templateProcessor,
            panel.GetType().GetField("_templateProcessor", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(panel));
    }

    [Test]
    public void TestCreateDataSourcePanel()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");
        var range = ws.Range(ws.Cell(1, 1), ws.Cell(2, 2));
        range.AddToNamed("d_Test", XLScope.Worksheet);
        var namedRange = ws.NamedRange("d_Test");

        var report = new object();
        var templateProcessor = Substitute.For<ITemplateProcessor>();
        var parseSettings = new PanelParsingSettings
        {
            DataSourcePanelPrefix = "d",
            PanelPrefixSeparator = "_"
        };

        var factory = new ExcelPanelFactory(report, templateProcessor, parseSettings);
        var panel = (ExcelDataSourcePanel) factory.Create(namedRange, new Dictionary<string, string>
        {
            [nameof(ExcelDataSourcePanel.Type)] = PanelType.Horizontal.ToString(),
            [nameof(ExcelDataSourcePanel.ShiftType)] = ShiftType.Row.ToString(),
            [nameof(ExcelDataSourcePanel.RenderPriority)] = "5",
            [nameof(ExcelDataSourcePanel.BeforeRenderMethodName)] = "BeforeRenderMethodName",
            [nameof(ExcelDataSourcePanel.AfterRenderMethodName)] = "AfterRenderMethodName",
            [nameof(ExcelDataSourcePanel.BeforeDataItemRenderMethodName)] = "BeforeDataItemRenderMethodName",
            [nameof(ExcelDataSourcePanel.AfterDataItemRenderMethodName)] = "AfterDataItemRenderMethodName",
            ["DataSource"] = "DS"
        });

        Assert.AreEqual(PanelType.Horizontal, panel.Type);
        Assert.AreEqual(ShiftType.Row, panel.ShiftType);
        Assert.AreEqual(5, panel.RenderPriority);
        Assert.AreEqual("BeforeRenderMethodName", panel.BeforeRenderMethodName);
        Assert.AreEqual("AfterRenderMethodName", panel.AfterRenderMethodName);
        Assert.AreEqual("BeforeDataItemRenderMethodName", panel.BeforeDataItemRenderMethodName);
        Assert.AreEqual("AfterDataItemRenderMethodName", panel.AfterDataItemRenderMethodName);
        Assert.AreEqual(0, panel.Children.Count);
        Assert.IsNull(panel.Parent);
        Assert.AreEqual(namedRange.Ranges.First(), panel.Range);
        Assert.AreSame("DS",
            panel.GetType().GetField("_dataSourceTemplate", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(panel));
        Assert.AreSame(report,
            panel.GetType().GetField("_report", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(panel));
        Assert.AreSame(templateProcessor,
            panel.GetType().GetField("_templateProcessor", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(panel));

        ExceptionAssert.Throws<InvalidOperationException>(() => factory.Create(namedRange, null),
            "Data source panel must have the property \"DataSource\"");
    }

    [Test]
    public void TestCreateDataSourceDynamicPanel()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");
        var range = ws.Range(ws.Cell(1, 1), ws.Cell(2, 2));
        range.AddToNamed("dyn_Test", XLScope.Worksheet);
        var namedRange = ws.NamedRange("dyn_Test");

        var report = new object();
        var templateProcessor = Substitute.For<ITemplateProcessor>();
        var parseSettings = new PanelParsingSettings
        {
            DynamicDataSourcePanelPrefix = "dyn",
            PanelPrefixSeparator = "_"
        };

        var factory = new ExcelPanelFactory(report, templateProcessor, parseSettings);
        var panel = (ExcelDataSourceDynamicPanel) factory.Create(namedRange, new Dictionary<string, string>
        {
            [nameof(ExcelDataSourceDynamicPanel.Type)] = PanelType.Horizontal.ToString(),
            [nameof(ExcelDataSourceDynamicPanel.ShiftType)] = ShiftType.Row.ToString(),
            [nameof(ExcelDataSourceDynamicPanel.RenderPriority)] = "5",
            [nameof(ExcelDataSourceDynamicPanel.BeforeRenderMethodName)] = "BeforeRenderMethodName",
            [nameof(ExcelDataSourceDynamicPanel.AfterRenderMethodName)] = "AfterRenderMethodName",
            [nameof(ExcelDataSourceDynamicPanel.BeforeDataItemRenderMethodName)] = "BeforeDataItemRenderMethodName",
            [nameof(ExcelDataSourceDynamicPanel.AfterDataItemRenderMethodName)] = "AfterDataItemRenderMethodName",
            [nameof(ExcelDataSourceDynamicPanel.BeforeHeadersRenderMethodName)] = "BeforeHeadersRenderMethodName",
            [nameof(ExcelDataSourceDynamicPanel.AfterHeadersRenderMethodName)] = "AfterHeadersRenderMethodName",
            [nameof(ExcelDataSourceDynamicPanel.BeforeDataTemplatesRenderMethodName)] =
                "BeforeDataTemplatesRenderMethodName",
            [nameof(ExcelDataSourceDynamicPanel.AfterDataTemplatesRenderMethodName)] =
                "AfterDataTemplatesRenderMethodName",
            [nameof(ExcelDataSourceDynamicPanel.BeforeDataRenderMethodName)] = "BeforeDataRenderMethodName",
            [nameof(ExcelDataSourceDynamicPanel.AfterDataRenderMethodName)] = "AfterDataRenderMethodName",
            [nameof(ExcelDataSourceDynamicPanel.BeforeTotalsTemplatesRenderMethodName)] =
                "BeforeTotalsTemplatesRenderMethodName",
            [nameof(ExcelDataSourceDynamicPanel.AfterTotalsTemplatesRenderMethodName)] =
                "AfterTotalsTemplatesRenderMethodName",
            [nameof(ExcelDataSourceDynamicPanel.BeforeTotalsRenderMethodName)] = "BeforeTotalsRenderMethodName",
            [nameof(ExcelDataSourceDynamicPanel.AfterTotalsRenderMethodName)] = "AfterTotalsRenderMethodName",
            ["DataSource"] = "DS"
        });

        Assert.AreEqual(PanelType.Horizontal, panel.Type);
        Assert.AreEqual(ShiftType.Row, panel.ShiftType);
        Assert.AreEqual(5, panel.RenderPriority);
        Assert.AreEqual("BeforeRenderMethodName", panel.BeforeRenderMethodName);
        Assert.AreEqual("AfterRenderMethodName", panel.AfterRenderMethodName);
        Assert.AreEqual("BeforeDataItemRenderMethodName", panel.BeforeDataItemRenderMethodName);
        Assert.AreEqual("AfterDataItemRenderMethodName", panel.AfterDataItemRenderMethodName);
        Assert.AreEqual("BeforeHeadersRenderMethodName", panel.BeforeHeadersRenderMethodName);
        Assert.AreEqual("AfterHeadersRenderMethodName", panel.AfterHeadersRenderMethodName);
        Assert.AreEqual("BeforeDataTemplatesRenderMethodName", panel.BeforeDataTemplatesRenderMethodName);
        Assert.AreEqual("AfterDataTemplatesRenderMethodName", panel.AfterDataTemplatesRenderMethodName);
        Assert.AreEqual("BeforeDataRenderMethodName", panel.BeforeDataRenderMethodName);
        Assert.AreEqual("AfterDataRenderMethodName", panel.AfterDataRenderMethodName);
        Assert.AreEqual("BeforeTotalsTemplatesRenderMethodName", panel.BeforeTotalsTemplatesRenderMethodName);
        Assert.AreEqual("AfterTotalsTemplatesRenderMethodName", panel.AfterTotalsTemplatesRenderMethodName);
        Assert.AreEqual("BeforeTotalsRenderMethodName", panel.BeforeTotalsRenderMethodName);
        Assert.AreEqual("AfterTotalsRenderMethodName", panel.AfterTotalsRenderMethodName);
        Assert.AreEqual(0, panel.Children.Count);
        Assert.IsNull(panel.Parent);
        Assert.AreEqual(namedRange.Ranges.First(), panel.Range);
        Assert.AreSame("DS",
            panel.GetType().GetField("_dataSourceTemplate", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(panel));
        Assert.AreSame(report,
            panel.GetType().GetField("_report", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(panel));
        Assert.AreSame(templateProcessor,
            panel.GetType().GetField("_templateProcessor", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(panel));

        ExceptionAssert.Throws<InvalidOperationException>(() => factory.Create(namedRange, null),
            "Dynamic data source panel must have the property \"DataSource\"");
    }

    [Test]
    public void TestCreateTotalsPanel()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");
        var range = ws.Range(ws.Cell(1, 1), ws.Cell(2, 2));
        range.AddToNamed("t_Test", XLScope.Worksheet);
        var namedRange = ws.NamedRange("t_Test");

        var report = new object();
        var templateProcessor = Substitute.For<ITemplateProcessor>();
        var parseSettings = new PanelParsingSettings
        {
            TotalsPanelPrefix = "t",
            PanelPrefixSeparator = "_"
        };

        var factory = new ExcelPanelFactory(report, templateProcessor, parseSettings);
        var panel = (ExcelTotalsPanel) factory.Create(namedRange, new Dictionary<string, string>
        {
            [nameof(ExcelTotalsPanel.Type)] = PanelType.Horizontal.ToString(),
            [nameof(ExcelTotalsPanel.ShiftType)] = ShiftType.Row.ToString(),
            [nameof(ExcelTotalsPanel.RenderPriority)] = "5",
            [nameof(ExcelTotalsPanel.BeforeRenderMethodName)] = "BeforeRenderMethodName",
            [nameof(ExcelTotalsPanel.AfterRenderMethodName)] = "AfterRenderMethodName",
            [nameof(ExcelTotalsPanel.BeforeDataItemRenderMethodName)] = "BeforeDataItemRenderMethodName",
            [nameof(ExcelTotalsPanel.AfterDataItemRenderMethodName)] = "AfterDataItemRenderMethodName",
            ["DataSource"] = "DS"
        });

        Assert.AreEqual(PanelType.Horizontal, panel.Type);
        Assert.AreEqual(ShiftType.Row, panel.ShiftType);
        Assert.AreEqual(5, panel.RenderPriority);
        Assert.AreEqual("BeforeRenderMethodName", panel.BeforeRenderMethodName);
        Assert.AreEqual("AfterRenderMethodName", panel.AfterRenderMethodName);
        Assert.AreEqual("BeforeDataItemRenderMethodName", panel.BeforeDataItemRenderMethodName);
        Assert.AreEqual("AfterDataItemRenderMethodName", panel.AfterDataItemRenderMethodName);
        Assert.AreEqual(0, panel.Children.Count);
        Assert.IsNull(panel.Parent);
        Assert.AreEqual(namedRange.Ranges.First(), panel.Range);
        Assert.AreSame("DS",
            panel.GetType().GetField("_dataSourceTemplate", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(panel));
        Assert.AreSame(report,
            panel.GetType().GetField("_report", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(panel));
        Assert.AreSame(templateProcessor,
            panel.GetType().GetField("_templateProcessor", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(panel));

        ExceptionAssert.Throws<InvalidOperationException>(() => factory.Create(namedRange, null),
            "Totals panel must have the property \"DataSource\"");
    }

    [Test]
    public void TestCreatePanelWithBadName()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");
        var range = ws.Range(ws.Cell(1, 1), ws.Cell(2, 2));
        range.AddToNamed("b_Test", XLScope.Worksheet);
        var namedRange = ws.NamedRange("b_Test");

        var report = new object();
        var templateProcessor = Substitute.For<ITemplateProcessor>();
        var parseSettings = new PanelParsingSettings {PanelPrefixSeparator = "-"};

        var factory = new ExcelPanelFactory(report, templateProcessor, parseSettings);
        ExceptionAssert.Throws<InvalidOperationException>(() => factory.Create(namedRange, null),
            "Panel name \"b_Test\" does not contain prefix separator \"-\"");
    }

    [Test]
    public void TestCreateUnsupportedPanelType()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");
        var range = ws.Range(ws.Cell(1, 1), ws.Cell(2, 2));
        range.AddToNamed("b_Test", XLScope.Worksheet);
        var namedRange = ws.NamedRange("b_Test");

        var report = new object();
        var templateProcessor = Substitute.For<ITemplateProcessor>();
        var parseSettings = new PanelParsingSettings {PanelPrefixSeparator = "_"};

        var factory = new ExcelPanelFactory(report, templateProcessor, parseSettings);
        ExceptionAssert.Throws<NotSupportedException>(() => factory.Create(namedRange, null),
            "Panel type with prefix \"b\" is not supported");
    }

    [Test]
    public void TestExcelPanelFactoryArgumentsCheck()
    {
        var report = new object();
        var templateProcessor = Substitute.For<ITemplateProcessor>();
        var parseSettings = new PanelParsingSettings();

        ExceptionAssert.Throws<ArgumentNullException>(() =>
            new ExcelPanelFactory(null, templateProcessor, parseSettings));
        ExceptionAssert.Throws<ArgumentNullException>(() => new ExcelPanelFactory(report, null, parseSettings));
        ExceptionAssert.Throws<ArgumentNullException>(() => new ExcelPanelFactory(report, templateProcessor, null));

        var factory = new ExcelPanelFactory(report, templateProcessor, parseSettings);
        ExceptionAssert.Throws<ArgumentNullException>(() => factory.Create(null, new Dictionary<string, string>()));
    }
}