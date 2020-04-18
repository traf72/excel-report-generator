using ClosedXML.Excel;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Rendering;
using ExcelReportGenerator.Rendering.EventArgs;
using ExcelReportGenerator.Rendering.Providers;
using ExcelReportGenerator.Rendering.Providers.DataItemValueProviders;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.Linq;
using System.Reflection;
using ExcelReportGenerator.Rendering.Providers.VariableProviders;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests
{
    public class TestReport : BaseReport
    {
        private int _counter;

        public TestReport()
        {
            ExpandoObj.StrProp = "ExpandoStr";
            ExpandoObj.DecimalProp = 5.56m;
        }

        public string StrParam { get; } = "String parameter";

        public int IntParam = 10;

        public DateTime DateParam { get; } = new DateTime(2017, 10, 25);

        public TimeSpan TimeSpanParam { get; set; } = TimeSpan.FromHours(20);

        public ComplexType ComplexTypeParam { get; set; } = new ComplexType();

        public dynamic ExpandoObj { get; set; } = new ExpandoObject();

        public string NullProp { get; set; }

        public string Format(DateTime date, string format = "yyyyMMdd")
        {
            return date.ToString(format);
        }

        public decimal Multiply(decimal num1, decimal num2)
        {
            return num1 * num2;
        }

        public string Concat(object item1, object item2)
        {
            return $"{item1}_{item2}";
        }

        public int Counter()
        {
            return ++_counter;
        }

        public void BeforeRenderParentDataSourcePanel(PanelBeforeRenderEventArgs args)
        {
            args.Range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        public void AfterRenderParentDataSourcePanelChildBottom(PanelEventArgs args)
        {
            args.Range.LastRow().Delete(XLShiftDeletedCells.ShiftCellsUp);
            args.Range.LastRow().Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            args.Range.LastRow().Style.Border.BottomBorderColor = XLColor.Black;
        }

        public void AfterRenderParentDataSourcePanelChildTop(PanelEventArgs args)
        {
            args.Range.FirstRow().Delete(XLShiftDeletedCells.ShiftCellsUp);
            args.Range.FirstRow().Style.Border.TopBorder = XLBorderStyleValues.Thin;
            args.Range.FirstRow().Style.Border.TopBorderColor = XLColor.Black;
        }

        public void AfterRenderParentDataSourcePanelChildRight(PanelEventArgs args)
        {
            args.Range.LastColumn().Delete(XLShiftDeletedCells.ShiftCellsLeft);
            args.Range.LastColumn().Style.Border.RightBorder = XLBorderStyleValues.Thin;
            args.Range.LastColumn().Style.Border.RightBorderColor = XLColor.Black;
        }

        public void AfterRenderParentDataSourcePanelChildLeft(PanelEventArgs args)
        {
            //// The standard way does not work. The range becomes Invalid (possible ClosedXml bug)
            //args.Range.FirstColumn().Delete(XLShiftDeletedCells.ShiftCellsLeft);
            //args.Range.FirstColumn().Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            //args.Range.FirstColumn().Style.Border.LeftBorderColor = XLColor.Black;

            IXLWorksheet worksheet = args.Range.Worksheet;
            IXLAddress firstColumnFirstCellAddress = args.Range.FirstColumn().FirstCell().Address;
            IXLAddress firstColumnLastCellAddress = args.Range.FirstColumn().LastCell().Address;

            args.Range.FirstColumn().Delete(XLShiftDeletedCells.ShiftCellsLeft);
            IXLRange range = worksheet.Range(firstColumnFirstCellAddress, firstColumnLastCellAddress);
            range.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            range.Style.Border.LeftBorderColor = XLColor.Black;
        }

        public void AfterRenderChildDataSourcePanel(PanelEventArgs args)
        {
            args.Range.LastRow().Delete(XLShiftDeletedCells.ShiftCellsUp);
        }

        public void CancelPanelRender(PanelBeforeRenderEventArgs args)
        {
            args.IsCanceled = true;
        }

        public void TestExcelPanelBeforeRender(PanelBeforeRenderEventArgs args)
        {
            args.Range.Cell(1, 1).Value = "{p:BoolParam}";
        }

        public void TestExcelPanelAfterRender(PanelEventArgs args)
        {
            args.Range.Cell(1, 2).Value = Convert.ToInt32(args.Range.Cell(1, 2).Value) + 1;
        }

        public void TestExcelDataSourcePanelBeforeRender(DataSourcePanelBeforeRenderEventArgs args)
        {
            IList<TestItem> data = ((IEnumerable<TestItem>)args.Data).ToList();
            data[2].Name = "ChangedName";
            args.Range.LastCell().Delete(XLShiftDeletedCells.ShiftCellsLeft);
        }

        public void TestExcelDataSourcePanelAfterRender(DataSourcePanelEventArgs args)
        {
            args.Range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        public void TestExcelDataItemPanelBeforeRender(DataItemPanelBeforeRenderEventArgs args)
        {
            var dataItem = (TestItem)args.DataItem.Value;
            if (dataItem.Name == "Test1")
            {
                dataItem.Name = "Test1_Changed_Before";
            }
        }

        public void TestExcelDataItemPanelAfterRender(DataItemPanelEventArgs args)
        {
            var dataItem = (TestItem)args.DataItem.Value;
            if (dataItem.Name == "Test2")
            {
                args.Range.FirstCell().Value = "Test2_Changed_After";
            }
        }

        public void TestExcelTotalsPanelBeforeRender(DataSourcePanelBeforeRenderEventArgs args)
        {
            IList<TestItem> data = ((IEnumerable<TestItem>)args.Data).ToList();
            data[0].Sum = 55.78m;
        }

        public void TestExcelTotalsPanelAfterRender(DataSourcePanelEventArgs args)
        {
            args.Range.FirstCell().Value = "Changed plain text";
        }

        public decimal CustomAggregation(decimal result, decimal currentValue, int itemNumber)
        {
            return (result + currentValue) / 2 + itemNumber;
        }

        public string PostAggregation(decimal result, int itemsCount)
        {
            return ((result + itemsCount) / 3).ToString("F3");
        }

        public double PostAggregationRound(double result, int itemsCount)
        {
            return Math.Round(result, 2);
        }

        public void TestExcelDynamicPanelBeforeHeadersRender(DataSourceDynamicPanelBeforeRenderEventArgs args)
        {
            args.Columns[0].Width = 30;
            args.Columns[0].AggregateFunction = AggregateFunction.Avg;
            args.Columns[1].AdjustToContent = true;
            args.Columns.Add(new ExcelDynamicColumn("DynamicAdded", typeof(decimal?), "Dynamic added") { Width = 20 });
        }

        public void TestExcelDynamicPanelAfterHeadersRender(DataSourceDynamicPanelEventArgs args)
        {
            args.Range.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            args.Range.Style.Border.OutsideBorderColor = XLColor.Red;
            args.Range.Style.Font.Bold = true;
        }

        public void TestExcelDynamicPanelBeforeNumbersRender(DataSourceDynamicPanelBeforeRenderEventArgs args)
        {
            args.Range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }

        public void TestExcelDynamicPanelAfterNumbersRender(DataSourceDynamicPanelEventArgs args)
        {
            args.Range.Style.Fill.BackgroundColor = XLColor.Gray;
            args.Range.Style.Font.FontColor = XLColor.White;
        }

        public void TestExcelDynamicPanelBeforeDataTemplatesRender(DataSourcePanelBeforeRenderEventArgs args)
        {
            args.Range.Style.Font.Underline = XLFontUnderlineValues.Single;
        }

        public void TestExcelDynamicPanelAfterDataTemplatesRender(DataSourceDynamicPanelEventArgs args)
        {
            args.Range.Cells().ElementAt(5).Style.NumberFormat.Format = "#,0.0";
            args.Range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            args.Range.Style.Border.OutsideBorderColor = XLColor.Black;
        }

        public void TestExcelDynamicPanelBeforeDataRender(DataSourcePanelBeforeRenderEventArgs args)
        {
            var dataSet = (DataSet)args.Data;
            DataTable dataTable = dataSet.Tables[0];
            dataTable.Rows[2]["Type"] = 1;
            dataTable.Columns.Add(new DataColumn("DynamicAdded", typeof(decimal)));
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                dataTable.Rows[i]["DynamicAdded"] = (i + 1) * 2.6;
            }
        }

        public void TestExcelDynamicPanelAfterDataRender(DataSourcePanelEventArgs args)
        {
            args.Range.Style.Border.InsideBorder = args.Range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            args.Range.Style.Border.InsideBorderColor = args.Range.Style.Border.OutsideBorderColor = XLColor.Orange;
        }

        public void TestExcelDynamicPanelBeforeDataItemRender(DataItemPanelBeforeRenderEventArgs args)
        {
            args.Range.LastCell().Style.Fill.BackgroundColor = XLColor.BlueGreen;
        }

        public void TestExcelDynamicPanelAfterDataItemRender(DataItemPanelEventArgs args)
        {
            IXLCell targetCell = args.Range.Cells().ElementAt(2);
            if (targetCell.Value is bool val)
            {
                targetCell.Value = val ? "Yes" : "No";
            }
        }

        public void TestExcelDynamicPanelBeforeTotalsTemplatesRender(DataSourcePanelBeforeRenderEventArgs args)
        {
            args.Range.Style.Font.Italic = true;
        }

        public void TestExcelDynamicPanelAfterTotalsTemplatesRender(DataSourcePanelEventArgs args)
        {
            args.Range.Style.Border.OutsideBorder = XLBorderStyleValues.Dashed;
            args.Range.Style.Border.OutsideBorderColor = XLColor.Green;
        }

        public void TestExcelDynamicPanelBeforeTotalsRender(DataSourcePanelBeforeRenderEventArgs args)
        {
            args.Range.Cells().ElementAt(1).Value = "{Count(di:Name)}";
        }

        public void TestExcelDynamicPaneAfterTotalsRender(DataSourcePanelEventArgs args)
        {
            args.Range.Cells().ElementAt(5).Style.NumberFormat.Format = "$ #,0.00";
        }

        public void TestExcelDynamicPaneBeforeRender(DataSourcePanelBeforeRenderEventArgs args)
        {
            args.Range.Cells().ElementAt(0).Value = "CanceledHeaders";
            args.Range.Cells().ElementAt(1).Value = "CanceledData";
            args.Range.Cells().ElementAt(2).Value = "CanceledTotals";
            args.IsCanceled = true;
        }

        public void TestExcelDynamicPaneAfterRender(DataSourcePanelEventArgs args)
        {
            args.Range.Style.Fill.BackgroundColor = XLColor.TractorRed;
        }
    }

    public class BaseReport
    {
        protected BaseReport()
        {
            Workbook = new XLWorkbook();
            var typeProvider = new DefaultTypeProvider(new[] { Assembly.GetExecutingAssembly(), Assembly.GetAssembly(typeof(DateTime)), }, GetType());
            var instanceProvider = new DefaultInstanceProvider(this);

            TemplateProcessor = new DefaultTemplateProcessor(new DefaultPropertyValueProvider(typeProvider, instanceProvider), new SystemVariableProvider(),
                new DefaultMethodCallValueProvider(typeProvider, instanceProvider), new DefaultDataItemValueProvider(new DataItemValueProviderFactory())
                {
                    DataItemSelfTemplate = "di"
                });
        }

        public bool BoolParam { get; set; } = true;

        public ITemplateProcessor TemplateProcessor { get; set; }

        public XLWorkbook Workbook { get; set; }

        public static DateTime AddDays(DateTime date, int daysCount)
        {
            return date.AddDays(daysCount);
        }
    }

    public class ComplexType
    {
        public string StrParam { get; set; } = "Complex type string parameter";

        public int IntParam = 11;
    }
}