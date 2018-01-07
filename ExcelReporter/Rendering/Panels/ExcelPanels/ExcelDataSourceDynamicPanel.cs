﻿using ClosedXML.Excel;
using ExcelReporter.Attributes;
using ExcelReporter.Enums;
using ExcelReporter.Excel;
using ExcelReporter.Extensions;
using ExcelReporter.Rendering.EventArgs;
using ExcelReporter.Rendering.Providers.ColumnsProviders;
using ExcelReporter.Rendering.TemplateProcessors;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelReporter.Rendering.Panels.ExcelPanels
{
    internal class ExcelDataSourceDynamicPanel : ExcelDataSourcePanel
    {
        private readonly IColumnsProviderFactory _columnsFactory = new ColumnsProviderFactory();

        public ExcelDataSourceDynamicPanel(string dataSourceTemplate, IXLNamedRange namedRange, object report, ITemplateProcessor templateProcessor)
            : base(dataSourceTemplate, namedRange, report, templateProcessor)
        {
        }

        public ExcelDataSourceDynamicPanel(object data, IXLNamedRange namedRange, object report, ITemplateProcessor templateProcessor)
            : base(data, namedRange, report, templateProcessor)
        {
        }

        [ExternalProperty]
        public string BeforeHeadersRenderMethodName { get; set; }

        [ExternalProperty]
        public string AfterHeadersRenderMethodName { get; set; }

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
            // Parent context does not affect on this panel type therefore don't care about it
            _data = _isDataReceivedDirectly ? _data : _templateProcessor.GetValue(_dataSourceTemplate);

            bool isCanceled = CallBeforeRenderMethod();
            if (isCanceled)
            {
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

            RenderHeaders(columns);
            IXLRange dataRange = RenderDataTemplates(columns);
            if (dataRange != null)
            {
                RenderData(dataRange);
            }

            IXLRange totalsRange = RenderTotalsTemplates(columns);
            if (totalsRange != null)
            {
                RenderTotals(totalsRange);
            }

            RemoveName();
            CallAfterRenderMethod();
        }

        private void RenderHeaders(IList<ExcelDynamicColumn> columns)
        {
            string template = _templateProcessor.WrapTemplate("Headers");
            IXLCell cell = Range.CellsUsed().SingleOrDefault(c => Regex.IsMatch(c.Value.ToString(), $@"^{template}$", RegexOptions.IgnoreCase));
            if (cell == null)
            {
                return;
            }

            IXLWorksheet ws = Range.Worksheet;
            IXLRange range = ws.Range(cell, cell);

            bool isCanceled = CallBeforeRenderMethod(BeforeHeadersRenderMethodName, range, columns);
            if (isCanceled)
            {
                return;
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

            IXLRange resultRange = GetColumnsRange(ws, cell, columns.Count);
            SetColumnsWidth(resultRange, columns);

            CallAfterRenderMethod(AfterHeadersRenderMethodName, resultRange, columns);
        }

        private IXLRange RenderDataTemplates(IList<ExcelDynamicColumn> columns)
        {
            string template = _templateProcessor.WrapTemplate("Data");
            IXLCell cell = Range.CellsUsed().SingleOrDefault(c => Regex.IsMatch(c.Value.ToString(), $@"^{template}$", RegexOptions.IgnoreCase));
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

            cell.Value = _templateProcessor.BuildDataItemTemplate("Template");
            string rangeName = $"DataTemplates_{Guid.NewGuid():N}";
            range.AddToNamed(rangeName, XLScope.Worksheet);

            var panel = new ExcelDataSourcePanel(columns.Select(c => new { Template = _templateProcessor.BuildDataItemTemplate(c.Name) }).ToList(),
                ws.NamedRange(rangeName), _report, _templateProcessor)
            {
                ShiftType = ShiftType.Cells,
                Type = Type == PanelType.Vertical ? PanelType.Horizontal : PanelType.Vertical,
            };

            panel.Render();

            IXLRange resultRange = GetColumnsRange(ws, cell, columns.Count);
            SetColumnsWidth(resultRange, columns);
            SetCellsDisplayFormat(resultRange, columns);

            CallAfterRenderMethod(AfterDataTemplatesRenderMethodName, resultRange, columns);

            return resultRange;
        }

        public void RenderData(IXLRange dataRange)
        {
            string rangeName = $"DynamicPanelData_{Guid.NewGuid():N}";
            dataRange.AddToNamed(rangeName, XLScope.Worksheet);
            var dataPanel = new ExcelDataSourcePanel(_data, Range.Worksheet.NamedRange(rangeName), _report, _templateProcessor)
            {
                ShiftType = ShiftType,
                Type = Type,
                BeforeRenderMethodName = BeforeDataRenderMethodName,
                AfterRenderMethodName = AfterDataRenderMethodName,
                BeforeDataItemRenderMethodName = BeforeDataItemRenderMethodName,
                AfterDataItemRenderMethodName = AfterDataItemRenderMethodName,
            };

            dataPanel.Render();
        }

        private IXLRange RenderTotalsTemplates(IList<ExcelDynamicColumn> columns)
        {
            string template = _templateProcessor.WrapTemplate("Totals");
            IXLCell cell = Range.CellsUsed().SingleOrDefault(c => Regex.IsMatch(c.Value.ToString(), $@"^{template}$", RegexOptions.IgnoreCase));
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

            cell.Value = _templateProcessor.BuildDataItemTemplate("Totals");
            string rangeName = $"Totals_{Guid.NewGuid():N}";
            range.AddToNamed(rangeName, XLScope.Worksheet);

            IList<string> totalsTemplates = new List<string>();
            foreach (ExcelDynamicColumn column in columns)
            {
                totalsTemplates.Add(column.AggregateFunction != AggregateFunction.NoAggregation
                    ? _templateProcessor.BuildAggregationFuncTemplate(column.AggregateFunction, column.Name)
                    : null);
            }

            var panel = new ExcelDataSourcePanel(totalsTemplates.Select(t => new { Totals = t }), ws.NamedRange(rangeName), _report, _templateProcessor)
            {
                ShiftType = ShiftType.Cells,
                Type = Type == PanelType.Vertical ? PanelType.Horizontal : PanelType.Vertical,
            };

            panel.Render();

            IXLRange resultRange = GetColumnsRange(ws, cell, columns.Count);
            SetColumnsWidth(resultRange, columns);
            SetCellsDisplayFormat(resultRange, columns);

            CallAfterRenderMethod(AfterTotalsTemplatesRenderMethodName, resultRange, columns);

            return resultRange;
        }

        public void RenderTotals(IXLRange totalsRange)
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
        }

        private IXLRange GetColumnsRange(IXLWorksheet ws, IXLCell rangeFirstCell, int columnsCount)
        {
            return Type == PanelType.Vertical
                ? ws.Range(rangeFirstCell, ExcelHelper.ShiftCell(rangeFirstCell, new AddressShift(0, columnsCount - 1)))
                : ws.Range(rangeFirstCell, ExcelHelper.ShiftCell(rangeFirstCell, new AddressShift(columnsCount - 1, 0)));
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

        //TODO Проверить корректное копирование, если передан не шаблон, а сами данные
        protected override IExcelPanel CopyPanel(IXLCell cell)
        {
            var panel = new ExcelDataSourceDynamicPanel(_dataSourceTemplate, CopyNamedRange(cell), _report, _templateProcessor);
            FillCopyProperties(panel);
            return panel;
        }
    }
}