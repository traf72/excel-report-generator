using ClosedXML.Excel;
using ExcelReporter.Enumerators;
using ExcelReporter.Enums;
using ExcelReporter.Excel;
using ExcelReporter.Helpers;
using ExcelReporter.Rendering.EventArgs;
using ExcelReporter.Reports;
using System;
using System.Collections;
using System.Linq;

namespace ExcelReporter.Rendering.Panels.ExcelPanels
{
    internal class ExcelDataSourcePanel : ExcelNamedPanel
    {
        protected readonly string _dataSourceTemplate;
        protected readonly bool _isDataReceivedDirectly;
        protected object _data;

        public ExcelDataSourcePanel(string dataSourceTemplate, IXLNamedRange namedRange, IExcelReport report)
            : base(namedRange, report)
        {
            if (string.IsNullOrWhiteSpace(dataSourceTemplate))
            {
                throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(dataSourceTemplate));
            }
            _dataSourceTemplate = dataSourceTemplate;
        }

        public ExcelDataSourcePanel(object data, IXLNamedRange namedRange, IExcelReport report) : base(namedRange, report)
        {
            _data = data ?? throw new ArgumentNullException(nameof(data), ArgumentHelper.NullParamMessage);
            _isDataReceivedDirectly = true;
        }

        public string BeforeDataItemRenderMethodName { get; set; }

        public string AfterDataItemRenderMethodName { get; set; }

        public override void Render()
        {
            // Получаем контекст родительского элемента данных, если он есть
            HierarchicalDataItem parentDataItem = GetDataContext();

            _data = _isDataReceivedDirectly ? _data : Report.TemplateProcessor.GetValue(_dataSourceTemplate, parentDataItem);

            bool isCanceled = CallBeforeRenderMethod();
            if (isCanceled)
            {
                return;
            }

            IEnumerator enumerator = null;
            try
            {
                enumerator = EnumeratorFactory.Create(_data);
                // Если данных нет, то просто удаляем сам шаблон
                if (enumerator == null || !enumerator.MoveNext())
                {
                    DeletePanel(this);
                    return;
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
                    currentPanel.Render();
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

            CallAfterRenderMethod();
        }

        private ExcelDataItemPanel CreateTemplatePanel()
        {
            var templatePanel = new ExcelDataItemPanel(Range, Report)
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

        protected override PanelEventArgs GetAfterPanelRenderEventArgs()
        {
            return new DataSourcePanelEventArgs { Range = Range, Data = _data };
        }

        //TODO Проверить корректное копирование, если передан не шаблон, а сами данные
        protected override IExcelPanel CopyPanel(IXLCell cell)
        {
            var panel = new ExcelDataSourcePanel(_dataSourceTemplate, CopyNamedRange(cell), Report)
            {
                BeforeDataItemRenderMethodName = BeforeDataItemRenderMethodName,
                AfterDataItemRenderMethodName = AfterDataItemRenderMethodName,
            };
            FillCopyProperties(panel);
            return panel;
        }
    }
}