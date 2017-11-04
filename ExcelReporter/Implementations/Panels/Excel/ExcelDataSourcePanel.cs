using ClosedXML.Excel;
using ExcelReporter.Enumerators;
using ExcelReporter.Enums;
using ExcelReporter.Excel;
using ExcelReporter.Interfaces.Panels.Excel;
using ExcelReporter.Interfaces.Reports;
using System.Collections;

namespace ExcelReporter.Implementations.Panels.Excel
{
    internal class ExcelDataSourcePanel : ExcelNamedPanel
    {
        private readonly string _dataSourceTemplate;

        public ExcelDataSourcePanel(string dataSourceTemplate, IXLNamedRange namedRange, IExcelReport report)
            : base(namedRange, report)
        {
            _dataSourceTemplate = dataSourceTemplate;
        }

        public override void Render()
        {
            // Получаем контекст родительского элемента данных, если он есть
            HierarchicalDataItem parentDataItem = GetDataContext();

            object data = Report.TemplateProcessor.GetValue(_dataSourceTemplate, parentDataItem);
            IEnumerator enumerator = EnumeratorFactory.Create(data);
            
            // Если данных нет, то просто удаляем сам шаблон
            if (enumerator == null || !enumerator.MoveNext())
            {
                DeletePanel(this);
                return;
            }

            object currentItem = enumerator.Current;
            var templatePanel = new ExcelDataItemPanel(Range, Report)
            {
                Parent = Parent,
                Children = Children,
                RenderPriority = RenderPriority,
                ShiftType = ShiftType,
                Type = Type,
            };
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

        private void DeletePanel(IExcelPanel panel)
        {
            RemoveAllNamesRecursive(panel);
            panel.Delete();
        }

        protected override IExcelPanel CopyPanel(IXLCell cell)
        {
            var panel = new ExcelDataSourcePanel(_dataSourceTemplate, CopyNamedRange(cell), Report);
            FillCopyProperties(panel);
            return panel;
        }
    }
}