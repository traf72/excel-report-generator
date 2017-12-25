using ClosedXML.Excel;
using ExcelReporter.Enums;
using ExcelReporter.Excel;
using ExcelReporter.Extensions;
using ExcelReporter.Rendering.Providers.ColumnsProviders;
using ExcelReporter.Reports;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelReporter.Rendering.Panels.ExcelPanels
{
    internal class ExcelDataSourceDynamicPanel : ExcelDataSourcePanel
    {
        private readonly IColumnsProviderFactory _columnsFactory = new ColumnsProviderFactory();

        public ExcelDataSourceDynamicPanel(string dataSourceTemplate, IXLNamedRange namedRange, IExcelReport report)
            : base(dataSourceTemplate, namedRange, report)
        {
        }

        public override void Render()
        {
            // Parent context does not affect on this panel type therefore don't care about it
            _data = Report.TemplateProcessor.GetValue(_dataSourceTemplate);
            IColumnsProvider columnsProvider = _columnsFactory.Create(_data);
            if (columnsProvider == null)
            {
                //TODO Обработать
                return;
            }

            IList<ExcelDynamicColumn> columns = columnsProvider.GetColumnsList(_data);
            if (!columns.Any())
            {
                //TODO Обработать
                return;
            }

            RenderHeaders(columns);
            IXLRange dataRange = RenderDataTemplates(columns);
            RenderData(dataRange);
            IXLRange totalsRange = RenderTotalsTemplates(columns);
            RenderTotals(totalsRange);
        }

        private void RenderHeaders(IList<ExcelDynamicColumn> columns)
        {
            string template = Report.TemplateProcessor.WrapTemplate("Headers");
            IXLCell cell = Range.CellsUsed().SingleOrDefault(c => Regex.IsMatch(c.Value.ToString(), $@"^{template}$", RegexOptions.IgnoreCase));
            if (cell == null)
            {
                return;
            }

            IXLWorksheet ws = Range.Worksheet;
            cell.Value = Report.TemplateProcessor.BuildDataItemTemplate(nameof(ExcelDynamicColumn.Caption));
            IXLRange range = ws.Range(cell, cell);
            string rangeName = $"Headers_{Guid.NewGuid():N}";
            range.AddToNamed(rangeName, XLScope.Worksheet);

            var panel = new ExcelDataSourcePanel(columns, ws.NamedRange(rangeName), Report)
            {
                ShiftType = ShiftType.Cells,
                Type = Type == PanelType.Vertical ? PanelType.Horizontal : PanelType.Vertical,
            };

            panel.Render();
        }

        private IXLRange RenderDataTemplates(IList<ExcelDynamicColumn> columns)
        {
            string dataTemplate = Report.TemplateProcessor.WrapTemplate("Data");
            IXLCell dataCell = Range.CellsUsed().SingleOrDefault(c => Regex.IsMatch(c.Value.ToString(), $@"^{dataTemplate}$", RegexOptions.IgnoreCase));
            if (dataCell == null)
            {
                return null;
            }

            IXLWorksheet ws = Range.Worksheet;
            dataCell.Value = Report.TemplateProcessor.BuildDataItemTemplate("Template");
            IXLRange dataTemplatesRange = ws.Range(dataCell, dataCell);
            string rangeName = $"DataTemplates_{Guid.NewGuid():N}";
            dataTemplatesRange.AddToNamed(rangeName, XLScope.Worksheet);

            var dataTemplatesPanel = new ExcelDataSourcePanel(columns.Select(c => new { Template = Report.TemplateProcessor.BuildDataItemTemplate(c.Name) }).ToList(),
                ws.NamedRange(rangeName), Report)
            {
                ShiftType = ShiftType.Cells,
                Type = Type == PanelType.Vertical ? PanelType.Horizontal : PanelType.Vertical,
            };

            dataTemplatesPanel.Render();
            return ws.Range(dataCell, ExcelHelper.ShiftCell(dataCell, new AddressShift(0, columns.Count)));
        }

        public void RenderData(IXLRange dataRange)
        {
            string rangeName = $"DynamicPanelData_{Guid.NewGuid():N}";
            dataRange.AddToNamed(rangeName, XLScope.Worksheet);
            var dataPanel = new ExcelDataSourcePanel(_data, Range.Worksheet.NamedRange(rangeName), Report)
            {
                ShiftType = ShiftType,
                Type = Type,
            };

            dataPanel.Render();
        }

        private IXLRange RenderTotalsTemplates(IList<ExcelDynamicColumn> columns)
        {
            string template = Report.TemplateProcessor.WrapTemplate("Totals");
            IXLCell cell = Range.CellsUsed().SingleOrDefault(c => Regex.IsMatch(c.Value.ToString(), $@"^{template}$", RegexOptions.IgnoreCase));
            if (cell == null)
            {
                return null;
            }

            IXLWorksheet ws = Range.Worksheet;
            cell.Value = Report.TemplateProcessor.BuildDataItemTemplate("Totals");
            IXLRange range = ws.Range(cell, cell);
            string rangeName = $"Totals_{Guid.NewGuid():N}";
            range.AddToNamed(rangeName, XLScope.Worksheet);

            IList<string> totalsTemplates = new List<string>();
            foreach (ExcelDynamicColumn column in columns)
            {
                totalsTemplates.Add(column.AggregateFunction != AggregateFunction.NoAggregation
                    ? Report.TemplateProcessor.BuildAggregationFuncTemplate(column.AggregateFunction, column.Name)
                    : null);
            }

            var panel = new ExcelDataSourcePanel(totalsTemplates.Select(t => new { Totals = t }), ws.NamedRange(rangeName), Report)
            {
                ShiftType = ShiftType.Cells,
                Type = Type == PanelType.Vertical ? PanelType.Horizontal : PanelType.Vertical,
            };

            panel.Render();
            return ws.Range(cell, ExcelHelper.ShiftCell(cell, new AddressShift(0, columns.Count)));
        }

        public void RenderTotals(IXLRange totalsRange)
        {
            string rangeName = $"DynamicPanelTotals_{Guid.NewGuid():N}";
            totalsRange.AddToNamed(rangeName, XLScope.Worksheet);
            _data = Report.TemplateProcessor.GetValue(_dataSourceTemplate);
            var totalsPanel = new ExcelTotalsPanel(_data, Range.Worksheet.NamedRange(rangeName), Report)
            {
                ShiftType = ShiftType,
                Type = Type,
            };

            totalsPanel.Render();
        }

        //TODO Проверить корректное копирование, если передан не шаблон, а сами данные
        protected override IExcelPanel CopyPanel(IXLCell cell)
        {
            var panel = new ExcelDataSourceDynamicPanel(_dataSourceTemplate, CopyNamedRange(cell), Report);
            FillCopyProperties(panel);
            return panel;
        }
    }
}