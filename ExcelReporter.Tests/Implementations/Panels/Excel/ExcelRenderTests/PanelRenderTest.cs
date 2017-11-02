using ClosedXML.Excel;
using ExcelReporter.Attributes;
using ExcelReporter.Implementations.Panels.Excel;
using ExcelReporter.Implementations.Providers;
using ExcelReporter.Implementations.Providers.DataItemValueProviders;
using ExcelReporter.Implementations.TemplateProcessors;
using ExcelReporter.Interfaces.Reports;
using ExcelReporter.Interfaces.TemplateProcessors;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using System.Reflection;

namespace ExcelReporter.Tests.Implementations.Panels.Excel.ExcelRenderTests
{
    [TestClass]
    public class PanelRenderTest
    {
        [TestMethod]
        public void TestRenderParameters()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(1, 1, 4, 5);

            ws.Cell(1, 1).Value = "{p:StrParam}";
            ws.Cell(1, 2).Value = "{p:IntParam}";
            ws.Cell(1, 3).Value = "{p:DateParam}";
            ws.Cell(1, 4).Value = "{p:BoolParam}";
            ws.Cell(1, 5).Value = "{p:TimeSpanParam}";
            ws.Cell(2, 1).Value = "{ p : StrParam }";
            ws.Cell(2, 2).Value = "Plain text";
            ws.Cell(2, 3).Value = "{Plain text}";
            ws.Cell(2, 4).Value = " { m : Format ( p : DateParam ) } ";
            ws.Cell(2, 5).Value = "''{m:Format(p:DateParam)}";
            ws.Cell(3, 1).Value = "Int: { p : IntParam }. Str: {p:StrParam}. FormattedDate: {m:Format(p:DateParam)}";
            ws.Cell(3, 2).Value = "''{m:Format(ms:AddDays(p:DateParam, p:IntParam), \"yyyy-MM-dd\")}";
            ws.Cell(3, 3).Value = "''{m:Format(ms:AddDays(p:DateParam, 5), ddMMyyyy)}";
            ws.Cell(3, 4).Value = "''{m:Format(ms:AddDays(p:DateParam, -2), dd.MM.yyyy)}";
            ws.Cell(3, 5).Value = "''{m:Format(ms:AddDays(p:DateParam, [int]-3), \"dd.MM.yyyy HH:mm:ss\")}";
            ws.Cell(4, 1).Value = "{m:TestReport:Counter()}";
            ws.Cell(4, 2).Value = "{ m : TestReport : Counter ( ) }";
            ws.Cell(4, 3).Value = "{m:Counter()}";
            ws.Cell(5, 1).Value = "{p:StrParam}";
            ws.Cell(5, 2).Value = "{m:Counter()}";
            ws.Cell(6, 1).Value = "Plain text outside range";

            var panel = new ExcelPanel(range, report);
            panel.Render();

            Assert.AreEqual(21, ws.CellsUsed().Count());
            Assert.AreEqual("String parameter", ws.Cell(1, 1).Value);
            Assert.AreEqual(10d, ws.Cell(1, 2).Value);
            Assert.AreEqual(new DateTime(2017, 10, 25), ws.Cell(1, 3).Value);
            Assert.AreEqual(true, ws.Cell(1, 4).Value);
            Assert.AreEqual(new TimeSpan(36500, 22, 30, 40), ws.Cell(1, 5).Value);
            Assert.AreEqual("String parameter", ws.Cell(2, 1).Value);
            Assert.AreEqual("Plain text", ws.Cell(2, 2).Value);
            Assert.AreEqual("{Plain text}", ws.Cell(2, 3).Value);
            Assert.AreEqual(20171025d, ws.Cell(2, 4).Value);
            Assert.AreEqual("20171025", ws.Cell(2, 5).Value);
            Assert.AreEqual("Int: 10. Str: String parameter. FormattedDate: 20171025", ws.Cell(3, 1).Value);
            Assert.AreEqual("2017-11-04", ws.Cell(3, 2).Value);
            Assert.AreEqual("30102017", ws.Cell(3, 3).Value);
            Assert.AreEqual("23.10.2017", ws.Cell(3, 4).Value);
            Assert.AreEqual("22.10.2017 00:00:00", ws.Cell(3, 5).Value);

            Assert.AreEqual(1d, ws.Cell(4, 1).Value);
            Assert.AreEqual(2d, ws.Cell(4, 2).Value);
            Assert.AreEqual(3d, ws.Cell(4, 3).Value);
            Assert.IsTrue(ws.Cell(4, 4).IsEmpty());
            Assert.IsTrue(ws.Cell(4, 5).IsEmpty());

            Assert.AreEqual("{p:StrParam}", ws.Cell(5, 1).Value);
            Assert.AreEqual("{m:Counter()}", ws.Cell(5, 2).Value);
            Assert.AreEqual("Plain text outside range", ws.Cell(6, 1).Value);

            //report.Workbook.SaveAs("test.xlsx");
        }

        private class TestReport : BaseReport
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

            public int Counter()
            {
                return ++_counter;
            }
        }

        private class BaseReport : IExcelReport
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
}