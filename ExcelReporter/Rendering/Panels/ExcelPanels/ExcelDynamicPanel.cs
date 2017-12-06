using ClosedXML.Excel;
using ExcelReporter.Enums;
using ExcelReporter.Excel;
using ExcelReporter.Extensions;
using ExcelReporter.Helpers;
using ExcelReporter.Rendering.Providers.ColumnsProviders;
using ExcelReporter.Reports;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelReporter.Rendering.Panels.ExcelPanels
{
    // Данная панель имеет жёсткую заточку на DefaultTemplateProcessor и ObjectPropertyValueProvider, то есть использование
    // данной панели невозможно при кастомной реализации интерфейса ITemplateProcessor - в будущем нужно устранить
    internal class ExcelDynamicPanel : ExcelNamedPanel
    {
        private readonly IColumnsProviderFactory _columnsFactory = new ColumnsProviderFactory();
        private readonly string _dataSourceTemplate;
        private object _data;

        public ExcelDynamicPanel(string dataSourceTemplate, IXLNamedRange namedRange, IExcelReport report)
            : base(namedRange, report)
        {
            if (string.IsNullOrWhiteSpace(dataSourceTemplate))
            {
                throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(dataSourceTemplate));
            }
            _dataSourceTemplate = dataSourceTemplate;
        }

        public ExcelDynamicPanel(object data, IXLNamedRange namedRange, IExcelReport report) : base(namedRange, report)
        {
            _data = data ?? throw new ArgumentNullException(nameof(data), ArgumentHelper.NullParamMessage);
        }

        public override void Render()
        {
            // Parent context does not affect on this panel type therefore don't care about it
            _data = _data ?? Report.TemplateProcessor.GetValue(_dataSourceTemplate);
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
            RenderTotals(columns);
        }

        private void RenderHeaders(IList<ExcelDynamicColumn> columns)
        {
            string template = Report.TemplateProcessor.WrapTemplate("Headers");
            IXLCell cell = Range.Cells().SingleOrDefault(c => Regex.IsMatch(c.Value.ToString(), $@"^{template}$", RegexOptions.IgnoreCase));
            if (cell == null)
            {
                return;
            }

            IXLWorksheet ws = Range.Worksheet;
            cell.Value = Report.TemplateProcessor.WrapTemplate($"di:{nameof(ExcelDynamicColumn.Caption)}");
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
            IXLCell dataCell = Range.Cells().SingleOrDefault(c => Regex.IsMatch(c.Value.ToString(), $@"^{dataTemplate}$", RegexOptions.IgnoreCase));
            if (dataCell == null)
            {
                return null;
            }

            IXLWorksheet ws = Range.Worksheet;
            dataCell.Value = Report.TemplateProcessor.WrapTemplate("di:di");
            IXLRange dataTemplatesRange = ws.Range(dataCell, dataCell);
            string rangeName = $"DataTemplates_{Guid.NewGuid():N}";
            dataTemplatesRange.AddToNamed(rangeName, XLScope.Worksheet);

            var dataTemplatesPanel = new ExcelDataSourcePanel(columns.Select(c => Report.TemplateProcessor.WrapTemplate($"di:{c.Name}")).ToList(), ws.NamedRange(rangeName), Report)
            {
                ShiftType = ShiftType.Cells,
                Type = Type == PanelType.Vertical ? PanelType.Horizontal : PanelType.Vertical,
            };

            dataTemplatesPanel.Render();

            IXLRange dataRange = ws.Range(dataCell, ExcelHelper.ShiftCell(dataCell, new AddressShift(0, columns.Count)));
            return dataRange;
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

        private void RenderTotals(IList<ExcelDynamicColumn> columns)
        {
            string template = Report.TemplateProcessor.WrapTemplate("Totals");
            IXLCell cell = Range.Cells().SingleOrDefault(c => Regex.IsMatch(c.Value.ToString(), $@"^{template}$", RegexOptions.IgnoreCase));
            if (cell == null)
            {
                return;
            }
            throw new NotImplementedException();
        }

        //TODO Проверить корректное копирование, если передан не шаблон, а сами данные
        protected override IExcelPanel CopyPanel(IXLCell cell)
        {
            var panel = new ExcelDynamicPanel(_dataSourceTemplate, CopyNamedRange(cell), Report);
            FillCopyProperties(panel);
            return panel;
        }
    }
}