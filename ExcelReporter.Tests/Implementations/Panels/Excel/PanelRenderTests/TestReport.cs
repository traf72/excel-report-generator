using ClosedXML.Excel;
using ExcelReporter.Attributes;
using ExcelReporter.Implementations.Providers;
using ExcelReporter.Implementations.Providers.DataItemValueProviders;
using ExcelReporter.Implementations.TemplateProcessors;
using ExcelReporter.Interfaces.Panels.Excel;
using ExcelReporter.Interfaces.Reports;
using ExcelReporter.Interfaces.TemplateProcessors;
using System;
using System.Reflection;

namespace ExcelReporter.Tests.Implementations.Panels.Excel.PanelRenderTests
{
    internal class TestReport : BaseReport
    {
        private int _counter;

        [Parameter]
        public string StrParam { get; } = "String parameter";

        [Parameter]
        public int IntParam { get; } = 10;

        [Parameter]
        public DateTime DateParam { get; } = new DateTime(2017, 10, 25);

        [Parameter]
        public TimeSpan TimeSpanParam { get; set; } = new TimeSpan(36500, 22, 30, 40);

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

        public void BeforeRenderParentDataSourcePanel(IExcelPanel panel)
        {
            panel.Range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        public void AfterRenderParentDataSourcePanelChildBottom(IExcelPanel panel)
        {
            panel.Range.LastRow().Delete(XLShiftDeletedCells.ShiftCellsUp);
            panel.Range.LastRow().Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            panel.Range.LastRow().Style.Border.BottomBorderColor = XLColor.Black;
        }

        public void AfterRenderParentDataSourcePanelChildTop(IExcelPanel panel)
        {
            panel.Range.FirstRow().Delete(XLShiftDeletedCells.ShiftCellsUp);
            panel.Range.FirstRow().Style.Border.TopBorder = XLBorderStyleValues.Thin;
            panel.Range.FirstRow().Style.Border.TopBorderColor = XLColor.Black;
        }

        public void AfterRenderParentDataSourcePanelChildRight(IExcelPanel panel)
        {
            panel.Range.LastColumn().Delete(XLShiftDeletedCells.ShiftCellsLeft);
            panel.Range.LastColumn().Style.Border.RightBorder = XLBorderStyleValues.Thin;
            panel.Range.LastColumn().Style.Border.RightBorderColor = XLColor.Black;
        }

        public void AfterRenderParentDataSourcePanelChildLeft(IExcelPanel panel)
        {
            panel.Range.FirstColumn().Delete(XLShiftDeletedCells.ShiftCellsLeft);
            panel.Range.FirstColumn().Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            panel.Range.FirstColumn().Style.Border.LeftBorderColor = XLColor.Black;
        }

        public void AfterRenderChildDataSourcePanel(IExcelPanel panel)
        {
            panel.Range.LastRow().Delete(XLShiftDeletedCells.ShiftCellsUp);
        }
    }

    internal class BaseReport : IExcelReport
    {
        protected BaseReport()
        {
            Workbook = new XLWorkbook();

            TemplateProcessor = new DefaultTemplateProcessor(new ReflectionParameterProvider(this),
                new MethodCallValueProvider(new TypeProvider(Assembly.GetExecutingAssembly()), this),
                new HierarchicalDataItemValueProvider(new DefaultDataItemValueProviderFactory()));
        }

        [Parameter]
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
}