using ClosedXML.Excel;
using ExcelReportGenerator.Attributes;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Excel;
using ExcelReportGenerator.Extensions;
using ExcelReportGenerator.Rendering.EventArgs;
using ExcelReportGenerator.Rendering.Providers.ColumnsProviders;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using System.Data;
using System.Text.RegularExpressions;

namespace ExcelReportGenerator.Rendering.Panels.ExcelPanels;

internal class ExcelDataSourceDynamicPanel : ExcelDataSourcePanel
{
    private readonly IColumnsProviderFactory _columnsFactory = new ColumnsProviderFactory();

    public ExcelDataSourceDynamicPanel(string dataSourceTemplate, IXLDefinedName namedRange, object report, ITemplateProcessor templateProcessor)
        : base(dataSourceTemplate, namedRange, report, templateProcessor)
    {
    }

    public ExcelDataSourceDynamicPanel(object data, IXLDefinedName namedRange, object report, ITemplateProcessor templateProcessor)
        : base(data, namedRange, report, templateProcessor)
    {
    }

    [ExternalProperty]
    public string BeforeHeadersRenderMethodName { get; set; }

    [ExternalProperty]
    public string AfterHeadersRenderMethodName { get; set; }

    [ExternalProperty]
    public string BeforeNumbersRenderMethodName { get; set; }

    [ExternalProperty]
    public string AfterNumbersRenderMethodName { get; set; }

    [ExternalProperty]
    public string BeforeDataTemplatesRenderMethodName { get; set; }

    [ExternalProperty]
    public string AfterDataTemplatesRenderMethodName { get; set; }

    [ExternalProperty]
    public string BeforeDataRenderMethodName { get; set; }

    [ExternalProperty]
    public string AfterDataRenderMethodName { get; set; }

    [ExternalProperty]
    public string BeforeTotalsTemplatesRenderMethodName { get; set; }

    [ExternalProperty]
    public string AfterTotalsTemplatesRenderMethodName { get; set; }

    [ExternalProperty]
    public string BeforeTotalsRenderMethodName { get; set; }

    [ExternalProperty]
    public string AfterTotalsRenderMethodName { get; set; }

    public override void Render()
    {
        // Receive parent data item context
        HierarchicalDataItem parentDataItem = GetDataContext();

        _data = _isDataReceivedDirectly ? _data : _templateProcessor.GetValue(_dataSourceTemplate, parentDataItem);

        bool isCanceled = CallBeforeRenderMethod();
        if (isCanceled)
        {
            ResultRange = ExcelHelper.CloneRange(Range);
            return;
        }

        IColumnsProvider columnsProvider = _columnsFactory.Create(_data);
        if (columnsProvider == null)
        {
            DeletePanel(this);
            return;
        }

        IList<ExcelDynamicColumn> columns = columnsProvider.GetColumnsList(_data);
        if (!columns.Any())
        {
            DeletePanel(this);
            return;
        }

        ResultRange = ExcelHelper.MergeRanges(Range, RenderHeaders(columns));
        ResultRange = ExcelHelper.MergeRanges(ResultRange, RenderColumnNumbers(columns));

        IXLRange dataRange = RenderDataTemplates(columns);
        if (dataRange != null)
        {
            ResultRange = ExcelHelper.MergeRanges(ResultRange, RenderData(dataRange));
        }

        IXLRange totalsRange = RenderTotalsTemplates(columns);
        if (totalsRange != null)
        {
            ResultRange = ExcelHelper.MergeRanges(ResultRange, RenderTotals(totalsRange));
        }

        RemoveName();
        CallAfterRenderMethod();
    }

    private IXLRange RenderHeaders(IList<ExcelDynamicColumn> columns)
    {
        string template = _templateProcessor.WrapTemplate("Headers");
        IXLCell cell = Range.CellsUsedWithoutFormulas().SingleOrDefault(c => Regex.IsMatch(c.Value.ToString(), $@"^{template}$", RegexOptions.IgnoreCase));
        if (cell == null)
        {
            return null;
        }

        IXLWorksheet ws = Range.Worksheet;
        IXLRange range = ws.Range(cell, cell);

        bool isCanceled = CallBeforeRenderMethod(BeforeHeadersRenderMethodName, range, columns);
        if (isCanceled)
        {
            return range;
        }

        cell.Value = _templateProcessor.BuildDataItemTemplate(nameof(ExcelDynamicColumn.Caption));
        string rangeName = $"Headers_{Guid.NewGuid():N}";
        range.AddToNamed(rangeName, XLScope.Worksheet);

        var panel = new ExcelDataSourcePanel(columns, ws.NamedRange(rangeName), _report, _templateProcessor)
        {
            ShiftType = ShiftType.Cells,
            Type = Type == PanelType.Vertical ? PanelType.Horizontal : PanelType.Vertical,
        };

        panel.Render();

        SetColumnsWidth(panel.ResultRange, columns);
        CallAfterRenderMethod(AfterHeadersRenderMethodName, panel.ResultRange, columns);

        return panel.ResultRange;
    }

    private IXLRange RenderColumnNumbers(IList<ExcelDynamicColumn> columns)
    {
        string template = _templateProcessor.WrapTemplate(@"Numbers(\((?<start>\d+)\))?");
        IXLCell cell = Range.CellsUsedWithoutFormulas().SingleOrDefault(c => Regex.IsMatch(c.Value.ToString(), $@"^{template}$", RegexOptions.IgnoreCase));
        if (cell == null)
        {
            return null;
        }

        IXLWorksheet ws = Range.Worksheet;
        IXLRange range = ws.Range(cell, cell);

        bool isCanceled = CallBeforeRenderMethod(BeforeNumbersRenderMethodName, range, columns);
        if (isCanceled)
        {
            return range;
        }

        Match match = Regex.Match(cell.Value.ToString(), $@"^{template}$", RegexOptions.IgnoreCase);
        if (!int.TryParse(match.Groups["start"]?.Value, out int startNumber))
        {
            startNumber = 1;
        }

        cell.Value = _templateProcessor.BuildDataItemTemplate(nameof(ColumnNumbersHelper.Number));
        string rangeName = $"ColumnNumbers_{Guid.NewGuid():N}";
        range.AddToNamed(rangeName, XLScope.Worksheet);

        var panel = new ExcelDataSourcePanel(columns.Select((c, i) => new ColumnNumbersHelper { Number = i + startNumber }).ToList(),
            ws.NamedRange(rangeName), _report, _templateProcessor)
        {
            ShiftType = ShiftType.Cells,
            Type = Type == PanelType.Vertical ? PanelType.Horizontal : PanelType.Vertical,
        };

        panel.Render();

        SetColumnsWidth(panel.ResultRange, columns);
        CallAfterRenderMethod(AfterNumbersRenderMethodName, panel.ResultRange, columns);

        return panel.ResultRange;
    }

    private IXLRange RenderDataTemplates(IList<ExcelDynamicColumn> columns)
    {
        string template = _templateProcessor.WrapTemplate("Data");
        IXLCell cell = Range.CellsUsedWithoutFormulas().SingleOrDefault(c => Regex.IsMatch(c.Value.ToString(), $@"^{template}$", RegexOptions.IgnoreCase));
        if (cell == null)
        {
            return null;
        }

        IXLWorksheet ws = Range.Worksheet;
        IXLRange range = ws.Range(cell, cell);

        bool isCanceled = CallBeforeRenderMethod(BeforeDataTemplatesRenderMethodName, range, columns);
        if (isCanceled)
        {
            return range;
        }

        cell.Value = _templateProcessor.BuildDataItemTemplate(nameof(DataTemplatesHelper.Template));
        string rangeName = $"DataTemplates_{Guid.NewGuid():N}";
        range.AddToNamed(rangeName, XLScope.Worksheet);

        var panel = new ExcelDataSourcePanel(columns.Select(c => new DataTemplatesHelper { Template = _templateProcessor.BuildDataItemTemplate(c.Name) }).ToList(),
            ws.NamedRange(rangeName), _report, _templateProcessor)
        {
            ShiftType = ShiftType.Cells,
            Type = Type == PanelType.Vertical ? PanelType.Horizontal : PanelType.Vertical,
        };

        panel.Render();

        SetColumnsWidth(panel.ResultRange, columns);
        SetCellsDisplayFormat(panel.ResultRange, columns);

        CallAfterRenderMethod(AfterDataTemplatesRenderMethodName, panel.ResultRange, columns);

        return panel.ResultRange;
    }

    public IXLRange RenderData(IXLRange dataRange)
    {
        string rangeName = $"DynamicPanelData_{Guid.NewGuid():N}";
        dataRange.AddToNamed(rangeName, XLScope.Worksheet);
        var dataPanel = new ExcelDataSourcePanel(_data, Range.Worksheet.NamedRange(rangeName), _report, _templateProcessor)
        {
            ShiftType = ShiftType,
            Type = Type,
            GroupBy = GroupBy,
            BeforeRenderMethodName = BeforeDataRenderMethodName,
            AfterRenderMethodName = AfterDataRenderMethodName,
            BeforeDataItemRenderMethodName = BeforeDataItemRenderMethodName,
            AfterDataItemRenderMethodName = AfterDataItemRenderMethodName,
        };

        dataPanel.Render();
        return dataPanel.ResultRange;
    }

    private IXLRange RenderTotalsTemplates(IList<ExcelDynamicColumn> columns)
    {
        string template = _templateProcessor.WrapTemplate("Totals");
        IXLCell cell = Range.CellsUsedWithoutFormulas().SingleOrDefault(c => Regex.IsMatch(c.Value.ToString(), $@"^{template}$", RegexOptions.IgnoreCase));
        if (cell == null)
        {
            return null;
        }

        IXLWorksheet ws = Range.Worksheet;
        IXLRange range = ws.Range(cell, cell);

        bool isCanceled = CallBeforeRenderMethod(BeforeTotalsTemplatesRenderMethodName, range, columns);
        if (isCanceled)
        {
            return range;
        }

        cell.Value = _templateProcessor.BuildDataItemTemplate(nameof(TotalsTemplatesHelper.Totals));
        string rangeName = $"Totals_{Guid.NewGuid():N}";
        range.AddToNamed(rangeName, XLScope.Worksheet);

        IList<string> totalsTemplates = new List<string>();
        foreach (ExcelDynamicColumn column in columns)
        {
            totalsTemplates.Add(column.AggregateFunction != AggregateFunction.NoAggregation
                ? _templateProcessor.BuildAggregationFuncTemplate(column.AggregateFunction, column.Name)
                : null);
        }

        var panel = new ExcelDataSourcePanel(totalsTemplates.Select(t => new TotalsTemplatesHelper { Totals = t }), ws.NamedRange(rangeName), _report, _templateProcessor)
        {
            ShiftType = ShiftType.Cells,
            Type = Type == PanelType.Vertical ? PanelType.Horizontal : PanelType.Vertical,
        };

        panel.Render();

        SetColumnsWidth(panel.ResultRange, columns);
        SetCellsDisplayFormat(panel.ResultRange, columns);

        CallAfterRenderMethod(AfterTotalsTemplatesRenderMethodName, panel.ResultRange, columns);

        return panel.ResultRange;
    }

    public IXLRange RenderTotals(IXLRange totalsRange)
    {
        string rangeName = $"DynamicPanelTotals_{Guid.NewGuid():N}";
        totalsRange.AddToNamed(rangeName, XLScope.Worksheet);

        if (_data is IDataReader dr && dr.IsClosed)
        {
            if (_isDataReceivedDirectly)
            {
                throw new InvalidOperationException("Cannot enumerate IDataReader twice. Cache data and try again.");
            }
            _data = _templateProcessor.GetValue(_dataSourceTemplate);
        }

        var totalsPanel = new ExcelTotalsPanel(_data, Range.Worksheet.NamedRange(rangeName), _report, _templateProcessor)
        {
            ShiftType = ShiftType,
            Type = Type,
            BeforeRenderMethodName = BeforeTotalsRenderMethodName,
            AfterRenderMethodName = AfterTotalsRenderMethodName,
        };

        totalsPanel.Render();
        return totalsPanel.ResultRange;
    }

    private bool CallBeforeRenderMethod(string methodName, IXLRange range, IList<ExcelDynamicColumn> columns)
    {
        if (string.IsNullOrWhiteSpace(methodName))
        {
            return false;
        }

        var args = new DataSourceDynamicPanelBeforeRenderEventArgs
        {
            Range = range,
            Columns = columns,
            Data = _data
        };

        CallReportMethod(methodName, new[] { args });
        return args.IsCanceled;
    }

    private void CallAfterRenderMethod(string methodName, IXLRange range, IList<ExcelDynamicColumn> columns)
    {
        if (string.IsNullOrWhiteSpace(methodName))
        {
            return;
        }

        var args = new DataSourceDynamicPanelEventArgs
        {
            Range = range,
            Columns = columns,
            Data = _data
        };

        CallReportMethod(methodName, new[] { args });
    }

    private void SetColumnsWidth(IXLRange range, IList<ExcelDynamicColumn> columns)
    {
        for (int i = 0; i < columns.Count; i++)
        {
            ExcelDynamicColumn column = columns[i];
            if (column.Width == null && !column.AdjustToContent)
            {
                continue;
            }

            if (Type == PanelType.Vertical)
            {
                IXLColumn excelColumn = range.Cell(1, i + 1).WorksheetColumn();
                if (column.Width != null)
                {
                    excelColumn.Width = column.Width.Value;
                }
                if (column.AdjustToContent)
                {
                    excelColumn.AdjustToContents();
                }
            }
            else
            {
                IXLRow excelRow = range.Cell(i + 1, 1).WorksheetRow();
                if (column.Width != null)
                {
                    excelRow.Height = column.Width.Value;
                }
                if (column.AdjustToContent)
                {
                    excelRow.AdjustToContents();
                }
            }
        }
    }

    private void SetCellsDisplayFormat(IXLRange range, IList<ExcelDynamicColumn> columns)
    {
        for (int i = 0; i < columns.Count; i++)
        {
            ExcelDynamicColumn column = columns[i];
            if (string.IsNullOrWhiteSpace(column.DisplayFormat) || column.DataType == null)
            {
                continue;
            }

            if (column.DataType.IsNumeric())
            {
                range.Cells().ElementAt(i).Style.NumberFormat.Format = column.DisplayFormat;
            }
            else if (column.DataType == typeof(DateTime) || column.DataType == typeof(DateTime?))
            {
                range.Cells().ElementAt(i).Style.DateFormat.Format = column.DisplayFormat;
            }
        }
    }

    protected override IExcelPanel CopyPanel(IXLCell cell)
    {
        var panel = _isDataReceivedDirectly
            ? new ExcelDataSourceDynamicPanel(_data, CopyNamedRange(cell), _report, _templateProcessor)
            : new ExcelDataSourceDynamicPanel(_dataSourceTemplate, CopyNamedRange(cell), _report, _templateProcessor);

        FillCopyProperties(panel);
        return panel;
    }

    protected override void FillCopyProperties(IExcelPanel panel)
    {
        var dataSourceDynamicPanel = panel as ExcelDataSourceDynamicPanel;
        dataSourceDynamicPanel.BeforeHeadersRenderMethodName = BeforeHeadersRenderMethodName;
        dataSourceDynamicPanel.AfterHeadersRenderMethodName = AfterHeadersRenderMethodName;
        dataSourceDynamicPanel.BeforeNumbersRenderMethodName = BeforeNumbersRenderMethodName;
        dataSourceDynamicPanel.AfterNumbersRenderMethodName = AfterNumbersRenderMethodName;
        dataSourceDynamicPanel.BeforeDataTemplatesRenderMethodName = BeforeDataTemplatesRenderMethodName;
        dataSourceDynamicPanel.AfterDataTemplatesRenderMethodName = AfterDataTemplatesRenderMethodName;
        dataSourceDynamicPanel.BeforeDataRenderMethodName = BeforeDataRenderMethodName;
        dataSourceDynamicPanel.AfterDataRenderMethodName = AfterDataRenderMethodName;
        dataSourceDynamicPanel.BeforeTotalsTemplatesRenderMethodName = BeforeTotalsTemplatesRenderMethodName;
        dataSourceDynamicPanel.AfterTotalsTemplatesRenderMethodName = AfterTotalsTemplatesRenderMethodName;
        dataSourceDynamicPanel.BeforeTotalsRenderMethodName = BeforeTotalsRenderMethodName;
        dataSourceDynamicPanel.AfterTotalsRenderMethodName = AfterTotalsRenderMethodName;

        base.FillCopyProperties(panel);
    }

    private class ColumnNumbersHelper
    {
        public int Number { get; set; }
    }

    private class DataTemplatesHelper
    {
        public string Template { get; set; }
    }

    private class TotalsTemplatesHelper
    {
        public string Totals { get; set; }
    }
}