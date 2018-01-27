using ClosedXML.Excel;
using ExcelReportGenerator.Enumerators;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Excel;
using ExcelReportGenerator.Helpers;
using ExcelReportGenerator.Rendering.EventArgs;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using System;
using System.Collections;
using System.Linq;
using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Rendering.Panels.ExcelPanels
{
    internal class ExcelDataSourcePanel : ExcelNamedPanel
    {
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

            IEnumerator enumerator = null;
            IXLRange resultRange = null;
            try
            {
                enumerator = EnumeratorFactory.Create(_data);
                // Если данных нет, то просто удаляем сам шаблон
                if (enumerator == null || !enumerator.MoveNext())
                {
                    DeletePanel(this);
                    return null;
                }

                object currentItem = enumerator.Current;
                // Создаём шаблон панели, который дальше будет размножаться
                var templatePanel = CreateTemplatePanel();
                while (true)
                {
                    ExcelDataItemPanel currentPanel;
                    bool nextItemExist = false;
                    object nextItem = null;
                    if (!enumerator.MoveNext())
                    {
                        // Если это последний элемент данных, то уже на размножаем шаблон, а рендерим данные напрямую в него
                        currentPanel = templatePanel;
                    }
                    else
                    {
                        nextItemExist = true;
                        nextItem = enumerator.Current;
                        // Сам шаблон сдвигаем вниз или вправо в зависимости от типа панели
                        ShiftTemplatePanel(templatePanel);
                        // Копируем шаблон на его предыдущее место
                        currentPanel = (ExcelDataItemPanel)templatePanel.Copy(ExcelHelper.ShiftCell(templatePanel.Range.FirstCell(), GetNextPanelAddressShift(templatePanel)));
                    }

                    currentPanel.DataItem = new HierarchicalDataItem { Value = currentItem, Parent = parentDataItem };
                    // Заполняем шаблон данными
                    IXLRange dataItemResultRange = currentPanel.Render();
                    resultRange = ExcelHelper.MergeRanges(resultRange, dataItemResultRange);
                    // Удаляем все сгенерированные имена Range'ей
                    RemoveAllNamesRecursive(currentPanel);

                    if (!nextItemExist)
                    {
                        break;
                    }

                    currentItem = nextItem;
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
                Children = Children.Select(c => c.Copy(c.Range.FirstCell())).ToList(),
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

        private AddressShift GetNextPanelAddressShift(IExcelPanel currentPanel)
        {
            return Type == PanelType.Vertical
                ? new AddressShift(-currentPanel.Range.RowCount(), 0)
                : new AddressShift(0, -currentPanel.Range.ColumnCount());
        }

        private void ShiftTemplatePanel(IExcelPanel templatePanel)
        {
            if (ShiftType == ShiftType.NoShift)
            {
                var addressShift = Type == PanelType.Vertical
                    ? new AddressShift(templatePanel.Range.RowCount(), 0)
                    : new AddressShift(0, templatePanel.Range.ColumnCount());

                templatePanel.Move(ExcelHelper.ShiftCell(templatePanel.Range.FirstCell(), addressShift));
            }
            else
            {
                ExcelHelper.AllocateSpaceForNextRange(templatePanel.Range, Type == PanelType.Vertical ? Direction.Top : Direction.Left, ShiftType);
            }
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