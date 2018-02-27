using ClosedXML.Excel;
using ExcelReportGenerator.Enumerators;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Extensions;
using ExcelReportGenerator.Helpers;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelReportGenerator.Excel;

namespace ExcelReportGenerator.Rendering.Panels.ExcelPanels
{
    internal class ExcelTotalsPanel : ExcelDataSourcePanel
    {
        public ExcelTotalsPanel(string dataSourceTemplate, IXLNamedRange namedRange, object report, ITemplateProcessor templateProcessor)
            : base(dataSourceTemplate, namedRange, report, templateProcessor)
        {
        }

        public ExcelTotalsPanel(object data, IXLNamedRange namedRange, object report, ITemplateProcessor templateProcessor)
            : base(data, namedRange, report, templateProcessor)
        {
        }

        public override void Render()
        {
            // Receieve parent data item context
            HierarchicalDataItem parentDataItem = GetDataContext();

            _data = _isDataReceivedDirectly ? _data : _templateProcessor.GetValue(_dataSourceTemplate, parentDataItem);

            bool isCanceled = CallBeforeRenderMethod();
            if (isCanceled)
            {
                ResultRange = Range;
                return;
            }

            ICustomEnumerator enumerator = null;
            try
            {
                enumerator = EnumeratorFactory.Create(_data) ?? new EnumerableEnumerator(new object[] { });
                IDictionary<IXLCell, IList<ParsedAggregationFunc>> totalCells = ParseTotalCells();
                DoAggregation(enumerator, totalCells.SelectMany(t => t.Value).ToArray(), parentDataItem);
                IXLWorksheet ws = Range.Worksheet;
                dynamic dataSource = new ExpandoObject();
                var dataSourceAsDict = (IDictionary<string, object>)dataSource;
                foreach (KeyValuePair<IXLCell, IList<ParsedAggregationFunc>> totalCell in totalCells)
                {
                    IList<ParsedAggregationFunc> aggFuncs = totalCell.Value;
                    foreach (ParsedAggregationFunc f in aggFuncs)
                    {
                        dataSourceAsDict[$"AggFunc_{f.UniqueName}"] = f.Result;
                    }
                }

                string rangeName = $"AggFuncs_{Guid.NewGuid():N}";
                Range.AddToNamed(rangeName, XLScope.Worksheet);

                var dataPanel = new ExcelDataSourcePanel(new[] { dataSource }, ws.NamedRange(rangeName), _report, _templateProcessor) { Parent = Parent };
                dataPanel.Render();
                ResultRange = ExcelHelper.MergeRanges(Range, dataPanel.ResultRange);
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
            string templatesWithAggregationRegexPattern = _templateProcessor.GetTemplatesWithAggregationRegexPattern();
            string aggregationFuncRegexPattern = _templateProcessor.GetAggregationFuncRegexPattern();
            foreach (IXLCell cell in Range.CellsUsedWithoutFormulas())
            {
                string cellValue = cell.Value.ToString();
                MatchCollection matches = Regex.Matches(cellValue, templatesWithAggregationRegexPattern, RegexOptions.IgnoreCase);
                if (matches.Count == 0)
                {
                    continue;
                }

                IList<ParsedAggregationFunc> aggFuncs = new List<ParsedAggregationFunc>();
                foreach (Match match in matches)
                {
                    string matchValue = match.Value;
                    MatchCollection innerMatches = Regex.Matches(match.Value, aggregationFuncRegexPattern, RegexOptions.IgnoreCase);
                    foreach (Match innerMatch in innerMatches)
                    {
                        string uniqueName = Guid.NewGuid().ToString("N");
                        matchValue = matchValue.ReplaceFirst(innerMatch.Value, _templateProcessor.UnwrapTemplate(_templateProcessor.BuildDataItemTemplate($"AggFunc_{uniqueName}")));
                        AggregateFunction aggFunc = EnumHelper.Parse<AggregateFunction>(innerMatch.Groups[1].Value);
                        string[] allFuncParams = new string[aggFuncMaxParamsCount];
                        string[] realFuncParams = innerMatch.Groups[2].Value.Trim().Split(',').Select(m => m.Trim()).ToArray();
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

                        aggFuncs.Add(new ParsedAggregationFunc(aggFunc, columnName.Trim()) { CustomFunc = allFuncParams[1], PostProcessFunction = allFuncParams[2], UniqueName = uniqueName });
                    }

                    cellValue = cellValue.ReplaceFirst(match.Value, matchValue);
                }

                cell.Value = cellValue;
                result[cell] = aggFuncs;
            }
            return result;
        }

        private void DoAggregation(IEnumerator enumerator, IList<ParsedAggregationFunc> aggFuncs, HierarchicalDataItem parentDataItem)
        {
            int dataItemsCount = 0;
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

                    dynamic value = _templateProcessor.GetValue(aggFunc.ColumnName, new HierarchicalDataItem { Value = item, Parent = parentDataItem });
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
            var panel = new ExcelTotalsPanel(_dataSourceTemplate, CopyNamedRange(cell), _report, _templateProcessor);
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

            /// <summary>
            /// Auxiliary property
            /// </summary>
            public string UniqueName { get; set; }
        }
    }
}