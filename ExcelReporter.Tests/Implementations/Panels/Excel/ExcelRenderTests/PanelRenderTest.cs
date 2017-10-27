using System;
using System.Reflection;
using ClosedXML.Excel;
using ExcelReporter.Attributes;
using ExcelReporter.Implementations.Panels.Excel;
using ExcelReporter.Implementations.Providers;
using ExcelReporter.Implementations.Providers.DataItemValueProviders;
using ExcelReporter.Implementations.TemplateProcessors;
using ExcelReporter.Interfaces.Reports;
using ExcelReporter.Interfaces.TemplateProcessors;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Implementations.Panels.Excel.ExcelRenderTests
{
    [TestClass]
    public class PanelRenderTest
    {
        [TestMethod]
        public void TestRenderParameters()
        {
        //    var report = new TestReport();
        //    IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
        //    IXLRange range = ws.Range(1, 1, 5, 5);

        //    ws.Cell(1, 1).Value = "{p:StrParam}";
        //    ws.Cell(1, 2).Value = "Plain text";
        //    ws.Cell(1, 3).DataType = XLCellValues.DateTime;
        //    ws.Cell(1, 3).Value = "'{p:DateParam}";
        //    ws.Cell(1, 4).Value = "''{p:DateParam}";
        //    //ws.Cell(1, 4).DataType = XLCellValues.Text;
        //    ws.Cell(1, 5).Value = "'{m:Format(p:DateParam)}";
        //    //ws.Cell(1, 5).DataType = XLCellValues.Text;
        //    ws.Cell(2, 1).Value = "{p:IntParam}";
        //    ws.Cell(2, 2).Value = "Int: {p:IntParam}. Date: {p:DateParam}. FormattedDate: {m:Format(p:DateParam)}";
        //    ws.Cell(7, 7).Value = "{p:StrParam}";

        //    var panel = new ExcelPanel(range, report);
        //    panel.Render();

        //    report.Workbook.SaveAs("test.xlsx");
        }

        private class TestReport : IExcelReport
        {
            public TestReport()
            {
                Workbook = new XLWorkbook();

                TemplateProcessor = new DefaultTemplateProcessor(new ReflectionParameterProvider(this),
                    new MethodCallValueProvider(new TypeProvider(Assembly.GetExecutingAssembly()), this),
                    new HierarchicalDataItemValueProvider(new DefaultDataItemValueProviderFactory()));
            }

            [Parameter]
            public string StrParam { get; } = "String parameter";

            [Parameter]
            public int IntParam { get; } = 100;

            [Parameter]
            public DateTime DateParam { get; } = new DateTime(2017, 10, 25);

            public ITemplateProcessor TemplateProcessor { get; set; }

            public XLWorkbook Workbook { get; set; }

            public string Format(DateTime date)
            {
                return date.ToString("yyyy-MMMM-dd");
            }

            public void Run()
            {
                throw new System.NotImplementedException();
            }
        }
    }
}