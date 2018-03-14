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
using System.Linq;
using ExcelReportGenerator.License;

namespace ExcelReportGenerator.Rendering.Panels.ExcelPanels
{
    internal class ExcelDataSourcePanel : ExcelNamedPanel
    {
        protected readonly string _dataSourceTemplate;
        protected readonly bool _isDataReceivedDirectly;
        protected object _data;

        private int _templatePanelRowCount;
        private int _templatePanelColumnCount;

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

        [System.Reflection.Obfuscation(Exclude = true, Feature = "renaming")]
        [ExternalProperty]
        public string GroupBy { get; set; }

        [System.Reflection.Obfuscation(Exclude = true, Feature = "renaming")]
        [ExternalProperty]
        public string BeforeDataItemRenderMethodName { get; set; }

        [System.Reflection.Obfuscation(Exclude = true, Feature = "renaming")]
        [ExternalProperty]
        public string AfterDataItemRenderMethodName { get; set; }

        public override void Render()
        {
            // Receieve parent data item context
            HierarchicalDataItem parentDataItem = GetDataContext();

            _data = _isDataReceivedDirectly ? _data : _templateProcessor.GetValue(_dataSourceTemplate, parentDataItem);

            bool isCanceled = CallBeforeRenderMethod();
            if (isCanceled)
            {
                ResultRange = ExcelHelper.CloneRange(Range);
                return;
            }

            ICustomEnumerator enumerator = null;
            try
            {
                enumerator = EnumeratorFactory.Create(_data);
                // Если данных нет, то просто удаляем сам шаблон
                if (enumerator == null || enumerator.RowCount == 0)
                {
                    DeletePanel(this);
                    return;
                }

                // Создаём шаблон панели, который дальше будет размножаться
                ExcelDataItemPanel templatePanel = CreateTemplatePanel();
                _templatePanelRowCount = templatePanel.Range.RowCount();
                _templatePanelColumnCount = templatePanel.Range.ColumnCount();

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
                    currentPanel.Render();
                    ResultRange = ExcelHelper.MergeRanges(ResultRange, currentPanel.ResultRange);

                    RemoveAllNamesRecursive(currentPanel);
                    rowNum++;
                }

                RemoveName();
            }
            finally
            {
                (enumerator as IDisposable)?.Dispose();
            }

            GroupResult();
            CallAfterRenderMethod();
        }

        private void GroupResult()
        {
            if (string.IsNullOrWhiteSpace(GroupBy))
            {
                return;
            }

            int[] groupColOrRowNumbers = GroupBy.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(strNum =>
            {
                if (int.TryParse(strNum, out int num))
                {
                    return num;
                }
                throw new InvalidCastException($"Parse \"{nameof(GroupBy)}\" property failed. Cannot convert value \"{strNum.Trim()}\" to {nameof(Int32)}");
            }).ToArray();

            if (Type == PanelType.Vertical)
            {
                GroupCellsVertical(ResultRange, groupColOrRowNumbers);
            }
            else
            {
                GroupCellsHorizontal(ResultRange, groupColOrRowNumbers);
            }
        }

        private void GroupCellsVertical(IXLRange range, int[] groupColNumbers)
        {
            IDictionary<int, (object StartCellValue, int StartRowNum)> previousCellValues = new Dictionary<int, (object, int)>();
            int rowsCount = range.Rows().Count();
            for (int rowNum = 1; rowNum <= rowsCount; rowNum++)
            {
                IXLRangeRow row = range.Row(rowNum);
                foreach (int colNum in groupColNumbers)
                {
                    object cellValue = row.Cell(colNum).Value;
                    if (previousCellValues.TryGetValue(colNum, out var previousResult))
                    {
                        if (!previousResult.StartCellValue.Equals(cellValue))
                        {
                            range.Range(previousResult.StartRowNum, colNum, rowNum - 1, colNum).Merge();
                            previousCellValues[colNum] = (cellValue, rowNum);
                        }
                        else if (rowNum == rowsCount)
                        {
                            range.Range(previousResult.StartRowNum, colNum, rowNum, colNum).Merge();
                        }
                    }
                    else
                    {
                        previousCellValues[colNum] = (cellValue, rowNum);
                    }
                }
            }
        }

        private void GroupCellsHorizontal(IXLRange range, int[] groupRowNumbers)
        {
            IDictionary<int, (object StartCellValue, int StartColNum)> previousCellValues = new Dictionary<int, (object, int)>();
            int colsCount = range.Columns().Count();
            for (int colNum = 1; colNum <= colsCount; colNum++)
            {
                IXLRangeColumn col = range.Column(colNum);
                foreach (int rowNum in groupRowNumbers)
                {
                    object cellValue = col.Cell(rowNum).Value;
                    if (previousCellValues.TryGetValue(rowNum, out var previousResult))
                    {
                        if (!previousResult.StartCellValue.Equals(cellValue))
                        {
                            range.Range(rowNum, previousResult.StartColNum, rowNum, colNum - 1).Merge();
                            previousCellValues[rowNum] = (cellValue, colNum);
                        }
                        else if (colNum == colsCount)
                        {
                            range.Range(rowNum, previousResult.StartColNum, rowNum, colNum).Merge();
                        }
                    }
                    else
                    {
                        previousCellValues[rowNum] = (cellValue, colNum);
                    }
                }
            }
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

            // For rarely call
            if (!(Parent is ExcelDataSourcePanel) && !(Parent is ExcelDataItemPanel))
            {
                // Check license
                if (Licensing.LicenseExpirationDate.Date < DateTime.Now.Date)
                {
                    throw new Exception(Licensing.LicenseViolationMessage);
                }
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
                ? new AddressShift(_templatePanelRowCount, 0)
                : new AddressShift(0, _templatePanelColumnCount);

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

        protected override PanelEventArgs GetAfterPanelRenderEventArgs()
        {
            return new DataSourcePanelEventArgs { Range = ResultRange, Data = _data };
        }

        //TODO Проверить корректное копирование, если передан не шаблон, а сами данные
        protected override IExcelPanel CopyPanel(IXLCell cell)
        {
            var panel = new ExcelDataSourcePanel(_dataSourceTemplate, CopyNamedRange(cell), _report, _templateProcessor)
            {
                GroupBy = GroupBy,
                BeforeDataItemRenderMethodName = BeforeDataItemRenderMethodName,
                AfterDataItemRenderMethodName = AfterDataItemRenderMethodName,
            };
            FillCopyProperties(panel);
            return panel;
        }
    }
}