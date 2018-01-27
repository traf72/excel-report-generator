using ClosedXML.Excel;
using ExcelReportGenerator.Enumerators;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Excel;
using ExcelReportGenerator.Helpers;
using ExcelReportGenerator.Rendering.EventArgs;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Rendering.Panels.ExcelPanels
{
    internal class ExcelDataSourcePanel : ExcelNamedPanel
    {
        public static IDictionary<string, TimeSpan> Benchmark = new Dictionary<string, TimeSpan>();

        protected readonly string _dataSourceTemplate;
        protected readonly bool _isDataReceivedDirectly;
        protected object _data;

        private int _templateRangeFirstRowOrColumnNumber;
        private int _templateRangeRowOrColumnCount;

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
            var totalSw = Stopwatch.StartNew();
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
                //if (Type == PanelType.Vertical)
                //{
                //    _templateRangeFirstRowOrColumnNumber = templatePanel.Range.FirstRow().RowNumber();
                //    _templateRangeRowOrColumnCount = templatePanel.Range.RowCount();
                //}
                //else
                //{
                //    _templateRangeFirstRowOrColumnNumber = templatePanel.Range.FirstColumn().ColumnNumber();
                //    _templateRangeRowOrColumnCount = templatePanel.Range.ColumnCount();
                //}

                var whileSw = Stopwatch.StartNew();

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
                        var shiftSw = Stopwatch.StartNew();

                        ShiftTemplatePanel(templatePanel);

                        shiftSw.Stop();
                        AddToBenchmark("Shift", shiftSw.Elapsed);


                        // Копируем шаблон на его предыдущее место
                        var copySw = Stopwatch.StartNew();

                        currentPanel = (ExcelDataItemPanel)templatePanel.Copy(ExcelHelper.ShiftCell(templatePanel.Range.FirstCell(), GetNextPanelAddressShift(templatePanel)));

                        copySw.Stop();
                        AddToBenchmark("Copy", copySw.Elapsed);
                    }

                    currentPanel.DataItem = new HierarchicalDataItem { Value = currentItem, Parent = parentDataItem };
                    // Заполняем шаблон данными

                    var renderSw = Stopwatch.StartNew();

                    IXLRange dataItemResultRange = currentPanel.Render();
                    //_templateRangeFirstRowOrColumnNumber++;
                    //if (dataItemResultRange != null && !dataItemResultRange.RangeAddress.IsInvalid)
                    //{
                    //    if (Type == PanelType.Vertical)
                    //    {
                    //        _templateRangeFirstRowOrColumnNumber += dataItemResultRange.RowCount();
                    //    }
                    //    else
                    //    {
                    //        _templateRangeFirstRowOrColumnNumber += dataItemResultRange.ColumnCount();
                    //    }
                    //}

                    renderSw.Stop();
                    AddToBenchmark("Render", renderSw.Elapsed);

                    var mergeSw = Stopwatch.StartNew();

                    resultRange = ExcelHelper.MergeRanges(resultRange, dataItemResultRange);

                    mergeSw.Stop();
                    AddToBenchmark("Merge", mergeSw.Elapsed);

                    // Удаляем все сгенерированные имена Range'ей

                    var removeSw = Stopwatch.StartNew();

                    RemoveAllNamesRecursive(currentPanel);

                    removeSw.Stop();
                    AddToBenchmark("Remove", removeSw.Elapsed);

                    if (!nextItemExist)
                    {
                        break;
                    }

                    currentItem = nextItem;
                }

                whileSw.Stop();
                AddToBenchmark("While", whileSw.Elapsed);

                RemoveName();
            }
            finally
            {
                (enumerator as IDisposable)?.Dispose();
            }

            CallAfterRenderMethod(resultRange);

            totalSw.Stop();
            AddToBenchmark("Total", totalSw.Elapsed);

            return resultRange;
        }

        private void AddToBenchmark(string key, TimeSpan elapsed)
        {
            if (Benchmark.TryGetValue(key, out TimeSpan current))
            {
                Benchmark[key] = current.Add(elapsed);
            }
            else
            {
                Benchmark[key] = elapsed;
            }
        }

        private ExcelDataItemPanel CreateTemplatePanel()
        {
            var tempWs = ExcelHelper.AddTempWorksheet(Range.Worksheet.Workbook);
            var tempRange = ExcelHelper.CopyRange(Range, tempWs.Cell(Range.FirstRow().RowNumber(), Range.FirstColumn().ColumnNumber()));

            //IDataReader dr = null;

            //SqlDataReader dr2 = null;

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
                //if (Type == PanelType.Vertical)
                //{
                //    if (ShiftType == ShiftType.Row)
                //    {
                //        templatePanel.Range.Worksheet.Row(_templateRangeFirstRowOrColumnNumber).InsertRowsAbove(_templateRangeRowOrColumnCount);
                //    }
                //}
                //else
                //{
                //    if (ShiftType == ShiftType.Row)
                //    {
                //        templatePanel.Range.Worksheet.Column(_templateRangeFirstRowOrColumnNumber).InsertColumnsBefore(_templateRangeRowOrColumnCount);
                //    }
                //}

                //templatePanel.Range.InsertRowsAbove(templatePanel.Range.RowCount(), true);
                //templatePanel.Range.InsertRowsBelow(templatePanel.Range.RowCount(), false);
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