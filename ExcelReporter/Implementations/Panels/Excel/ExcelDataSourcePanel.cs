using ClosedXML.Excel;
using ExcelReporter.Enums;
using ExcelReporter.Excel;
using ExcelReporter.Interfaces.Panels.Excel;
using ExcelReporter.Interfaces.Reports;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

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
            HierarchicalDataItem parentDataItem = GetDataContext();
            object data = Report.TemplateProcessor.GetValue(_dataSourceTemplate, parentDataItem);
            if (data is IEnumerable)
            {
                IList<object> listData = (data as IEnumerable).Cast<object>().ToList();
                // Если данных нет, то просто удаляем сам шаблон
                if (!listData.Any())
                {
                    DeletePanel(this);
                }
                else
                {
                    var templatePanel = new ExcelDataItemPanel(Range, Report)
                    {
                        Parent = Parent,
                        Children = Children,
                        RenderPriority = RenderPriority,
                        ShiftType = ShiftType,
                        Type = Type,
                    };

                    for (int i = 0; i < listData.Count; i++)
                    {
                        ExcelDataItemPanel currentPanel;
                        if (i != listData.Count - 1)
                        {
                            // Сам шаблон сдвигаем вниз или вправо в зависимости от типа панели
                            ShiftTemplatePanel(templatePanel);
                            // Копируем шаблон на его предыдущее место
                            currentPanel = (ExcelDataItemPanel)templatePanel.Copy(ExcelHelper.ShiftCell(templatePanel.Range.FirstCell(), GetNextPanelAddressShift(templatePanel)));
                        }
                        // Если это последний элемент данных, то уже на размножаем шаблон, а рендерим данные напрямую в него
                        else
                        {
                            currentPanel = templatePanel;
                        }

                        currentPanel.DataItem = new HierarchicalDataItem { Value = listData[i], Parent = parentDataItem };
                        // Заполняем шаблон данными
                        currentPanel.Render();
                        // Удаляем все сгенерированные имена Range'ей
                        RemoveAllNamesRecursive(currentPanel);
                    }
                    // Удаляем имя самого шаблона
                    RemoveName();
                }
            }
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