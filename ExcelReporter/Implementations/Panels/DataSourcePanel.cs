using ClosedXML.Excel;
using ExcelReporter.Enums;
using ExcelReporter.Excel;
using ExcelReporter.Interfaces.Panels;
using ExcelReporter.Interfaces.Reports;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace ExcelReporter.Implementations.Panels
{
    public class DataSourcePanel : NamedPanel, IDataSourcePanel
    {
        private readonly string _dataSourceName;
        private readonly string _dataSourceNethod;

        public DataSourcePanel(string dataSourceName, string dataSourceNethod, IXLNamedRange namedRange, IExcelReport report)
            : base(namedRange, report)
        {
            _dataSourceName = dataSourceName;
            _dataSourceNethod = dataSourceNethod;
        }

        public object Data { get; private set; }

        public override void Render()
        {
            object dataSource = GetDataSourceInstance();
            object result = CallDataSourceMethod(dataSource);
            if (result is IEnumerable)
            {
                IList<object> data = (result as IEnumerable).Cast<object>().ToList();
                // Если данных нет, то просто удаляем сам шаблон
                if (!data.Any())
                {
                    DeletePanel(this);
                }
                else
                {
                    // Создаем панель-шаблон для одного элемента данных
                    IPanel templatePanel = new Panel(Range, null)
                    {
                        Parent = Parent,
                        Children = Children,
                    };

                    for (int i = 0; i < data.Count; i++)
                    {
                        IPanel currentPanel;
                        if (i != data.Count - 1)
                        {
                            // Сам шаблон сдвигаем вниз или вправо в зависимости от типа панели
                            ShiftTemplatePanel(templatePanel);
                            // Копируем шаблон на его предыдущее место
                            currentPanel = templatePanel.Copy(ExcelHelper.ShiftCell(templatePanel.Range.FirstCell(), GetNextPanelAddressShift(templatePanel)));
                        }
                        // Если это последний элемент данных, то уже на размножаем шаблон, а рендерим данные напрямую в него
                        else
                        {
                            currentPanel = templatePanel;
                        }

                        //currentPanel.Report.TemplateProcessor.DataItemValueProvider = new DataItemValueProvider(data[i]);
                        // Заполняем шаблон данными
                        currentPanel.Render();
                        // Удаляем все сгенерированные имена Range'ей
                        RemoveAllNamesRecursive(currentPanel);
                    }
                    // Удаляем имя самого шаблона
                    RemoveName();
                }
            }

            ////Data = func.Invoke(dataSource, callParameters.ToArray());
        }

        private object GetDataSourceInstance()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            // TODO Пока простейший вариант поиска. Тут нужно будет подумать как правильно искать
            // с учётом того, что основной DataSource может быть переопределён для компании,
            // а также что могут быть общие DataSource'ы
            // Возможно сделать какой-то DataSourceProvider
            Type type = assembly.GetTypes().Single(t => t.Name == _dataSourceName);
            return Activator.CreateInstance(type);
        }

        private object CallDataSourceMethod(object dataSource)
        {
            Match match = Regex.Match(_dataSourceNethod, @"(.+)\((.*)\)");
            string methodName = match.Groups[1].Value;
            string methodParams = match.Groups[2].Value;
            object[] callParameters = methodParams
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(p => Report.TemplateProcessor.GetValue(p.Trim()))
                .ToArray();
            MethodInfo method = dataSource.GetType().GetMethod(methodName);
            return method.Invoke(dataSource, callParameters);
        }

        private AddressShift GetNextPanelAddressShift(IPanel currentPanel)
        {
            return Type == PanelType.Vertical
                ? new AddressShift(-currentPanel.Range.RowCount(), 0)
                : new AddressShift(0, -currentPanel.Range.ColumnCount());
        }

        private void ShiftTemplatePanel(IPanel templatePanel)
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

        private void DeletePanel(IPanel panel)
        {
            RemoveAllNamesRecursive(panel);
            panel.Delete();
        }

        protected override IPanel CopyPanel(IXLCell cell)
        {
            var panel = new DataSourcePanel(_dataSourceName, _dataSourceNethod, CopyNamedRange(cell), Report);
            FillCopyProperties(panel);
            return panel;
        }
    }
}