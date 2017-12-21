using ClosedXML.Excel;
using ExcelReporter.Rendering.Panels.ExcelPanels;
using ExcelReporter.Rendering.Providers;
using ExcelReporter.Rendering.Providers.DataItemValueProviders;
using ExcelReporter.Rendering.TemplateProcessors;
using ExcelReporter.Reports;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelReporter.Rendering.EventArgs;

namespace ExcelReporter.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests
{
    public class TestReport : BaseReport
    {
        private int _counter;

        public string StrParam { get; } = "String parameter";

        public int IntParam = 10;

        public DateTime DateParam { get; } = new DateTime(2017, 10, 25);

        public TimeSpan TimeSpanParam { get; set; } = new TimeSpan(36500, 22, 30, 40);

        public ComplexType ComplexTypeParam { get; set; } = new ComplexType();

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
            //// Стандартный способ не работает, Range почему-то становится Invalid (возможно баг ClosedXml)
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
            IList<TestItem> data = ((IEnumerable<TestItem>) args.Data).ToList();
            data[2].Name = "ChangedName";
        }

        public void TestExcelDataSourcePanelAfterRender(DataSourcePanelEventArgs args)
        {
            args.Range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        public void TestExcelDataItemPanelBeforeRender(DataItemPanelBeforeRenderEventArgs args)
        {
            var dataItem = (TestItem) args.DataItem.Value;
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
    }

    public class BaseReport : IExcelReport
    {
        protected BaseReport()
        {
            Workbook = new XLWorkbook();
            var typeProvider = new DefaultTypeProvider(new[] { Assembly.GetExecutingAssembly(), Assembly.GetAssembly(typeof(DateTime)), }, GetType());
            var instanceProvider = new DefaultInstanceProvider(this);

            TemplateProcessor = new DefaultTemplateProcessor(new DefaultPropertyValueProvider(typeProvider, instanceProvider),
                new DefaultMethodCallValueProvider(typeProvider, instanceProvider),
                new HierarchicalDataItemValueProvider(new DataItemValueProviderFactory()));
        }

        public bool BoolParam { get; set; } = true;

        public ITemplateProcessor TemplateProcessor { get; set; }

        public XLWorkbook Workbook { get; set; }

        public static DateTime AddDays(DateTime date, int daysCount)
        {
            return date.AddDays(daysCount);
        }

        public void Run()
        {
            throw new NotImplementedException();
        }
    }

    public class ComplexType
    {
        public string StrParam { get; set; } = "Complex type string parameter";

        public int IntParam = 11;
    }
}