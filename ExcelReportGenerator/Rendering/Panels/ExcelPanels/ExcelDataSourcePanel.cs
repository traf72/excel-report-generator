using ClosedXML.Excel;
using ExcelReportGenerator.Attributes;
using ExcelReportGenerator.Enumerators;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Excel;
using ExcelReportGenerator.Helpers;
using ExcelReportGenerator.Rendering.EventArgs;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using System;
using System.Collections.Generic;

namespace ExcelReportGenerator.Rendering.Panels.ExcelPanels
{
    internal class ExcelDataSourcePanel : ExcelNamedPanel
    {
        public static IDictionary<string, TimeSpan> Benchmark = new Dictionary<string, TimeSpan>();

        protected readonly string _dataSourceTemplate;
        protected readonly bool _isDataReceivedDirectly;
        protected object _data;

        public ExcelDataSourcePanel(string dataSourceTemplate, IXLNamedRange namedRange, object report, ITemplateProcessor templateProcessor)
            : base(namedRange, report, templateProcessor)
        {
            if (string.IsNullOrWhiteSpace(dataSourceTemplate))
            {
                throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(dataSourceTemplate));
            }
            _dataSourceTemplate = dataSourceTemplate;
        }

        public ExcelDataSourcePanel(object data, IXLNamedRange namedRange, object report, ITemplateProcessor templateProcessor) : base(namedRange, report, templateProcessor)
        {
            _data = data ?? throw new ArgumentNullException(nameof(data), ArgumentHelper.NullParamMessage);
            _isDataReceivedDirectly = true;
        }

        [ExternalProperty]
        public string BeforeDataItemRenderMethodName { get; set; }

        [ExternalProperty]
        public string AfterDataItemRenderMethodName { get; set; }

        public override IXLRange Render()
        {
            // Receieve parent data item context
            HierarchicalDataItem parentDataItem = GetDataContext();

            _data = _isDataReceivedDirectly ? _data : _templateProcessor.GetValue(_dataSourceTemplate, parentDataItem);

            bool isCanceled = CallBeforeRenderMethod();
            if (isCanceled)
            {
                return Range;
            }

            ICustomEnumerator enumerator = null;
            IXLRange resultRange = null;
            try
            {
                enumerator = EnumeratorFactory.Create(_data);
                // Если данных нет, то просто удаляем сам шаблон
                if (enumerator == null || enumerator.RowCount == 0)
                {
                    DeletePanel(this);
                    return null;
                }

                // Создаём шаблон панели, который дальше будет размножаться
                ExcelDataItemPanel templatePanel = CreateTemplatePanel();

                // Выделяем место под данные
                if (enumerator.RowCount > 1)
                {
                    AllocateSpaceForData(templatePanel, enumerator.RowCount);
                }

                int rowNum = 0;
                while (enumerator.MoveNext())
                {
                    object currentItem = enumerator.Current;
                    ExcelDataItemPanel currentPanel;
                    if (rowNum != enumerator.RowCount - 1)
                    {
                        IXLCell templateFirstCell = templatePanel.Range.FirstCell();
                        // Сам шаблон перемещаем вниз или вправо в зависимости от типа панели
                        MoveTemplatePanel(templatePanel);
                        // Копируем шаблон на его предыдущее место для панели, в которую будем ренедрить текущий элемент данных
                        currentPanel = (ExcelDataItemPanel)templatePanel.Copy(templateFirstCell);
                    }
                    else
                    {
                        // Если это последний элемент данных, то уже на размножаем шаблон, а рендерим данные напрямую в него
                        currentPanel = templatePanel;
                    }

                    currentPanel.DataItem = new HierarchicalDataItem { Value = currentItem, Parent = parentDataItem };

                    // Заполняем шаблон данными
                    IXLRange dataItemResultRange = currentPanel.Render();
                    resultRange = ExcelHelper.MergeRanges(resultRange, dataItemResultRange);

                    RemoveAllNamesRecursive(currentPanel);
                    rowNum++;
                }

                RemoveName();
            }
            finally
            {
                (enumerator as IDisposable)?.Dispose();
            }

            CallAfterRenderMethod(resultRange);
            return resultRange;
        }

        private ExcelDataItemPanel CreateTemplatePanel()
        {
            var templatePanel = new ExcelDataItemPanel(Range, _report, _templateProcessor)
            {
                Parent = Parent,
                Children = Children,
                RenderPriority = RenderPriority,
                ShiftType = ShiftType,
                Type = Type,
                BeforeRenderMethodName = BeforeDataItemRenderMethodName,
                AfterRenderMethodName = AfterDataItemRenderMethodName,
            };

            foreach (IExcelPanel child in templatePanel.Children)
            {
                child.Parent = templatePanel;
            }

            return templatePanel;
        }

        private void AllocateSpaceForData(IExcelPanel templatePanel, int dataItemsCount)
        {
            if (ShiftType == ShiftType.NoShift)
            {
                return;
            }

            IXLRange range = templatePanel.Range;
            if (Type == PanelType.Vertical)
            {
                int rowCount = (dataItemsCount - 1) * Range.RowCount();
                if (ShiftType == ShiftType.Row)
                {
                    range.Worksheet.Row(range.LastRow().RowNumber()).InsertRowsBelow(rowCount);
                }
                else
                {
                    range.InsertRowsBelow(rowCount, false);
                }
            }
            else
            {
                int columnCount = (dataItemsCount - 1) * Range.ColumnCount();
                if (ShiftType == ShiftType.Row)
                {
                    range.Worksheet.Column(range.LastColumn().ColumnNumber()).InsertColumnsAfter(columnCount);
                }
                else
                {
                    range.InsertColumnsAfter(columnCount, false);
                }
            }
        }

        private void MoveTemplatePanel(IExcelPanel templatePanel)
        {
            AddressShift shift = Type == PanelType.Vertical
                ? new AddressShift(templatePanel.Range.RowCount(), 0)
                : new AddressShift(0, templatePanel.Range.ColumnCount());

            templatePanel.Move(ExcelHelper.ShiftCell(templatePanel.Range.FirstCell(), shift));
        }

        protected void DeletePanel(IExcelPanel panel)
        {
            RemoveAllNamesRecursive(panel);
            panel.Delete();
        }

        protected override PanelBeforeRenderEventArgs GetBeforePanelRenderEventArgs()
        {
            return new DataSourcePanelBeforeRenderEventArgs { Range = Range, Data = _data };
        }

        protected override PanelEventArgs GetAfterPanelRenderEventArgs(IXLRange resultRange)
        {
            return new DataSourcePanelEventArgs { Range = resultRange, Data = _data };
        }

        //TODO Проверить корректное копирование, если передан не шаблон, а сами данные
        protected override IExcelPanel CopyPanel(IXLCell cell)
        {
            var panel = new ExcelDataSourcePanel(_dataSourceTemplate, CopyNamedRange(cell), _report, _templateProcessor)
            {
                BeforeDataItemRenderMethodName = BeforeDataItemRenderMethodName,
                AfterDataItemRenderMethodName = AfterDataItemRenderMethodName,
            };
            FillCopyProperties(panel);
            return panel;
        }
    }
}