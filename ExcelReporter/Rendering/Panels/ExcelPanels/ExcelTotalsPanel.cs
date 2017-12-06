using ClosedXML.Excel;
using ExcelReporter.Enumerators;
using ExcelReporter.Enums;
using ExcelReporter.Helpers;
using ExcelReporter.Rendering.Providers.DataItemValueProviders;
using ExcelReporter.Reports;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ExcelReporter.Rendering.Panels.ExcelPanels
{
    internal class ExcelTotalsPanel : ExcelNamedPanel
    {
        private readonly IDataItemValueProviderFactory _dataItemValueProviderFactory = new DataItemValueProviderFactory();
        private readonly string _dataSourceTemplate;
        private object _data;

        public ExcelTotalsPanel(string dataSourceTemplate, IXLNamedRange namedRange, IExcelReport report) : base(namedRange, report)
        {
            if (string.IsNullOrWhiteSpace(dataSourceTemplate))
            {
                throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(dataSourceTemplate));
            }
            _dataSourceTemplate = dataSourceTemplate;
        }

        internal ExcelTotalsPanel(object data, IXLNamedRange namedRange, IExcelReport report) : base(namedRange, report)
        {
            _data = data ?? throw new ArgumentNullException(nameof(data), ArgumentHelper.NullParamMessage);
        }

        public override void Render()
        {
            // Parent context does not affect on this panel type therefore don't care about it
            _data = _data ?? Report.TemplateProcessor.GetValue(_dataSourceTemplate);
            IEnumerator enumerator = null;
            try
            {
                enumerator = EnumeratorFactory.Create(_data);
                // Если данных нет, то просто удаляем сам шаблон, а может быть нужно скинуть всё в 0?
                if (enumerator == null)
                {
                    Delete();
                    return;
                }

                IList<TotalCellInfo> totalCells = ParseTotalCells();
                DoAggregation(enumerator, totalCells);
            }
            finally
            {
                (enumerator as IDisposable)?.Dispose();
            }
        }

        private IList<TotalCellInfo> ParseTotalCells()
        {
            IList<TotalCellInfo> result = new List<TotalCellInfo>();
            foreach (IXLCell cell in Range.CellsUsed())
            {
                string[] aggFuncs = Enum.GetNames(typeof(AggregateFunction));
                Match match = Regex.Match(cell.Value.ToString(), $@"^({string.Join("|", aggFuncs)}):(.+)$", RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    // TODO доработать парсинг для случая CustomAggregationFunc, а также для PostProcessFunc
                    AggregateFunction aggFunc = EnumHelper.Parse<AggregateFunction>(match.Groups[1].Value);
                    string columnName = match.Groups[2].Value;
                    result.Add(new TotalCellInfo(cell, aggFunc, columnName));
                }
            }
            return result;
        }

        private void DoAggregation(IEnumerator enumerator, IList<TotalCellInfo> totalCells)
        {
            int dataItemsCount = 0;
            IDataItemValueProvider valueProvider = null;
            while (enumerator.MoveNext())
            {
                dataItemsCount++;
                object item = enumerator.Current;
                foreach (TotalCellInfo totalCell in totalCells)
                {
                    if (totalCell.AggregateFunction == AggregateFunction.Count)
                    {
                        totalCell.Result = dataItemsCount;
                        continue;
                    }

                    valueProvider = valueProvider ?? _dataItemValueProviderFactory.Create(item);
                    dynamic value = valueProvider.GetValue(totalCell.ColumnName, item);
                    if (totalCell.AggregateFunction == AggregateFunction.Custom)
                    {
                        if (string.IsNullOrWhiteSpace(totalCell.CustomFunc))
                        {
                            throw new InvalidOperationException("The custom type of aggregation is specified in the template but custom function is missing");
                        }
                        totalCell.Result = CallReportMethod(totalCell.CustomFunc, new object[] { totalCell.Result, value, dataItemsCount });
                        continue;
                    }

                    if (value == null || value.Equals(DBNull.Value))
                    {
                        continue;
                    }
                    if (totalCell.Result == null)
                    {
                        totalCell.Result = value;
                        continue;
                    }

                    switch (totalCell.AggregateFunction)
                    {
                        case AggregateFunction.Sum:
                        case AggregateFunction.Avg:
                            totalCell.Result += value;
                            break;

                        case AggregateFunction.Min:
                        case AggregateFunction.Max:
                            var comparable = totalCell.Result as IComparable;
                            if (comparable != null && totalCell.Result.GetType() == value.GetType())
                            {
                                int compareResult = comparable.CompareTo(value);
                                if (totalCell.AggregateFunction == AggregateFunction.Min)
                                {
                                    totalCell.Result = compareResult < 0 ? totalCell.Result : value;
                                }
                                else
                                {
                                    totalCell.Result = compareResult < 0 ? value : totalCell.Result;
                                }
                            }
                            else
                            {
                                throw new InvalidOperationException("For Min and Max aggregation functions data items must implement IComparable interface");
                            }
                            break;

                        default:
                            throw new NotSupportedException("Unsupportable aggregation function");
                    }
                }
            }

            foreach (TotalCellInfo totalCell in totalCells)
            {
                if (totalCell.Result == null && (totalCell.AggregateFunction == AggregateFunction.Sum
                    || totalCell.AggregateFunction == AggregateFunction.Avg
                    || totalCell.AggregateFunction == AggregateFunction.Count))
                {
                    totalCell.Result = 0;
                }

                if (dataItemsCount != 0 && totalCell.AggregateFunction == AggregateFunction.Avg)
                {
                    totalCell.Result = (double)totalCell.Result / dataItemsCount;
                }

                if (!string.IsNullOrWhiteSpace(totalCell.PostProcessFunction))
                {
                    totalCell.Result = CallReportMethod(totalCell.PostProcessFunction, new[] { totalCell.Result, dataItemsCount });
                }
            }
        }

        //TODO Проверить корректное копирование, если передан не шаблон, а сами данные
        protected override IExcelPanel CopyPanel(IXLCell cell)
        {
            var panel = new ExcelTotalsPanel(_dataSourceTemplate, CopyNamedRange(cell), Report);
            FillCopyProperties(panel);
            return panel;
        }

        internal class TotalCellInfo
        {
            public TotalCellInfo(IXLCell cell, AggregateFunction aggregateFunction, string columnName)
            {
                Cell = cell;
                AggregateFunction = aggregateFunction;
                ColumnName = columnName;
            }

            public IXLCell Cell { get; }

            public AggregateFunction AggregateFunction { get; }

            public string ColumnName { get; }

            /// <summary>
            /// Call if AggregateFunction == AggregateFunction.Custom
            /// </summary>
            public string CustomFunc { get; set; }

            /// <summary>
            /// Call when aggregation completed
            /// </summary>
            public string PostProcessFunction { get; set; }

            public dynamic Result { get; set; }
        }
    }
}