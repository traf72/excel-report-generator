using ClosedXML.Excel;
using ExcelReporter.Enumerators;
using ExcelReporter.Enums;
using ExcelReporter.Extensions;
using ExcelReporter.Helpers;
using ExcelReporter.Rendering.Providers.DataItemValueProviders;
using ExcelReporter.Reports;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelReporter.Rendering.Panels.ExcelPanels
{
    internal class ExcelTotalsPanel : ExcelDataSourcePanel
    {
        private readonly IDataItemValueProviderFactory _dataItemValueProviderFactory = new DataItemValueProviderFactory();

        public ExcelTotalsPanel(string dataSourceTemplate, IXLNamedRange namedRange, IExcelReport report) : base(dataSourceTemplate, namedRange, report)
        {
        }

        public ExcelTotalsPanel(object data, IXLNamedRange namedRange, IExcelReport report) : base(data, namedRange, report)
        {
        }

        public override void Render()
        {
            // Parent context does not affect on this panel type therefore don't care about it
            _data = _data ?? Report.TemplateProcessor.GetValue(_dataSourceTemplate);

            bool isCanceled = CallBeforeRenderMethod();
            if (isCanceled)
            {
                return;
            }

            IEnumerator enumerator = null;
            try
            {
                enumerator = EnumeratorFactory.Create(_data) ?? Enumerable.Empty<object>().GetEnumerator();
                IDictionary<IXLCell, IList<ParsedAggregationFunc>> totalCells = ParseTotalCells();
                DoAggregation(enumerator, totalCells.SelectMany(t => t.Value).ToArray());
                foreach (KeyValuePair<IXLCell, IList<ParsedAggregationFunc>> totalCell in totalCells)
                {
                    IXLCell cell = totalCell.Key;
                    IList<ParsedAggregationFunc> aggFuncs = totalCell.Value;
                    cell.Value = cell.Value.ToString() == "{0}"
                        ? aggFuncs.First().Result
                        : string.Format(cell.Value.ToString(), aggFuncs.Select(f => f.Result).ToArray());
                }
            }
            finally
            {
                (enumerator as IDisposable)?.Dispose();
            }

            RemoveName();
            CallAfterRenderMethod();
        }

        private IDictionary<IXLCell, IList<ParsedAggregationFunc>> ParseTotalCells()
        {
            var result = new Dictionary<IXLCell, IList<ParsedAggregationFunc>>();
            const int aggFuncMaxParamsCount = 3;
            string aggregationRegexPattern = Report.TemplateProcessor.GetFullAggregationRegexPattern();
            foreach (IXLCell cell in Range.CellsUsed())
            {
                string cellValue = cell.Value.ToString();
                MatchCollection matches = Regex.Matches(cellValue, aggregationRegexPattern, RegexOptions.IgnoreCase);
                if (matches.Count == 0)
                {
                    continue;
                }

                IList<ParsedAggregationFunc> aggFuncs = new List<ParsedAggregationFunc>();
                for (int i = 0; i < matches.Count; i++)
                {
                    Match match = matches[i];
                    cellValue = cellValue.ReplaceFirst(match.Value, $"{{{i}}}");
                    AggregateFunction aggFunc = EnumHelper.Parse<AggregateFunction>(match.Groups[1].Value);
                    string[] allFuncParams = new string[aggFuncMaxParamsCount];
                    string[] realFuncParams = match.Groups[2].Value.Trim().Split(',').Select(m => m.Trim()).ToArray();
                    if (!realFuncParams.Any() || realFuncParams.Length > aggFuncMaxParamsCount)
                    {
                        throw new InvalidOperationException($"Aggregation function must have at least one but no more than {aggFuncMaxParamsCount} parameters");
                    }

                    for (int j = 0; j < realFuncParams.Length; j++)
                    {
                        allFuncParams[j] = string.IsNullOrWhiteSpace(realFuncParams[j]) ? null : realFuncParams[j];
                    }

                    string columnName = allFuncParams[0];
                    if (columnName == null)
                    {
                        throw new InvalidOperationException("\"ColumnName\" parameter in aggregation function cannot be empty");
                    }

                    columnName = Report.TemplateProcessor.TrimDataItemLabel(columnName);
                    aggFuncs.Add(new ParsedAggregationFunc(aggFunc, columnName) { CustomFunc = allFuncParams[1], PostProcessFunction = allFuncParams[2] });
                }
                cell.Value = cellValue;
                result[cell] = aggFuncs;
            }
            return result;
        }

        private void DoAggregation(IEnumerator enumerator, IList<ParsedAggregationFunc> aggFuncs)
        {
            int dataItemsCount = 0;
            IDataItemValueProvider valueProvider = null;
            while (enumerator.MoveNext())
            {
                dataItemsCount++;
                object item = enumerator.Current;
                foreach (ParsedAggregationFunc aggFunc in aggFuncs)
                {
                    if (aggFunc.AggregateFunction == AggregateFunction.Count)
                    {
                        aggFunc.Result = dataItemsCount;
                        continue;
                    }

                    valueProvider = valueProvider ?? _dataItemValueProviderFactory.Create(item);
                    dynamic value = valueProvider.GetValue(aggFunc.ColumnName, item);
                    if (aggFunc.AggregateFunction == AggregateFunction.Custom)
                    {
                        if (string.IsNullOrWhiteSpace(aggFunc.CustomFunc))
                        {
                            throw new InvalidOperationException("The custom type of aggregation is specified in the template but custom function is missing");
                        }
                        aggFunc.Result = CallReportMethod(aggFunc.CustomFunc, new object[] { aggFunc.Result, value, dataItemsCount });
                        continue;
                    }

                    if (value == null || value.Equals(DBNull.Value))
                    {
                        continue;
                    }
                    if (aggFunc.Result == null)
                    {
                        aggFunc.Result = value;
                        continue;
                    }

                    switch (aggFunc.AggregateFunction)
                    {
                        case AggregateFunction.Sum:
                        case AggregateFunction.Avg:
                            aggFunc.Result += value;
                            break;

                        case AggregateFunction.Min:
                        case AggregateFunction.Max:
                            var comparable = aggFunc.Result as IComparable;
                            if (comparable != null && aggFunc.Result.GetType() == value.GetType())
                            {
                                int compareResult = comparable.CompareTo(value);
                                if (aggFunc.AggregateFunction == AggregateFunction.Min)
                                {
                                    aggFunc.Result = compareResult < 0 ? aggFunc.Result : value;
                                }
                                else
                                {
                                    aggFunc.Result = compareResult < 0 ? value : aggFunc.Result;
                                }
                            }
                            else
                            {
                                throw new InvalidOperationException($"For {nameof(AggregateFunction.Min)} and {nameof(AggregateFunction.Max)} aggregation functions data items must implement IComparable interface");
                            }
                            break;

                        default:
                            throw new NotSupportedException("Unsupportable aggregation function");
                    }
                }
            }

            foreach (ParsedAggregationFunc aggFunc in aggFuncs)
            {
                if (aggFunc.Result == null && (aggFunc.AggregateFunction == AggregateFunction.Sum
                    || aggFunc.AggregateFunction == AggregateFunction.Avg
                    || aggFunc.AggregateFunction == AggregateFunction.Count))
                {
                    aggFunc.Result = 0;
                }

                if (dataItemsCount != 0 && aggFunc.AggregateFunction == AggregateFunction.Avg)
                {
                    aggFunc.Result = (double)aggFunc.Result / dataItemsCount;
                }

                if (!string.IsNullOrWhiteSpace(aggFunc.PostProcessFunction))
                {
                    aggFunc.Result = CallReportMethod(aggFunc.PostProcessFunction, new object[] { aggFunc.Result, dataItemsCount });
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

        internal class ParsedAggregationFunc
        {
            public ParsedAggregationFunc(AggregateFunction aggregateFunction, string columnName)
            {
                AggregateFunction = aggregateFunction;
                ColumnName = columnName;
            }

            public AggregateFunction AggregateFunction { get; }

            public string ColumnName { get; }

            /// <summary>
            /// Call if AggregateFunction = AggregateFunction.Custom
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