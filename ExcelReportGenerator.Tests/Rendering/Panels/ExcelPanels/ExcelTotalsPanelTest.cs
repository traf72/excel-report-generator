using ClosedXML.Excel;
using ExcelReportGenerator.Enumerators;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Exceptions;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using ExcelReportGenerator.Tests.CustomAsserts;
using ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using DataTable = System.Data.DataTable;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels
{
    [TestClass]
    public class ExcelTotalsPanelTest
    {
        [TestMethod]
        public void TestDoAggregation()
        {
            var dataTable = new DataTable();
            dataTable.Columns.Add(new DataColumn("TestColumn1", typeof(int)));
            dataTable.Columns.Add(new DataColumn("TestColumn2", typeof(decimal)));
            dataTable.Columns.Add(new DataColumn("TestColumn3", typeof(string)));
            dataTable.Columns.Add(new DataColumn("TestColumn4", typeof(bool)));
            dataTable.Rows.Add(3, 20.7m, "abc", false);
            dataTable.Rows.Add(1, 10.5m, "jkl", true);
            dataTable.Rows.Add(null, null, null, null);
            dataTable.Rows.Add(2, 30.9m, "def", false);

            var totalPanel = new ExcelTotalsPanel(dataTable, Substitute.For<IXLNamedRange>(), Substitute.For<object>(), Substitute.For<ITemplateProcessor>());
            IEnumerator enumerator = EnumeratorFactory.Create(dataTable);
            IList<ExcelTotalsPanel.ParsedAggregationFunc> totalCells = new List<ExcelTotalsPanel.ParsedAggregationFunc>
            {
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Sum, "TestColumn1"),
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Sum, "TestColumn2"),
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Sum, "TestColumn3"),
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Count, "TestColumn1"),
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Count, "TestColumn3"),
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Avg, "TestColumn1"),
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Avg, "TestColumn2"),
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Min, "TestColumn1"),
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Max, "TestColumn1"),
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Min, "TestColumn2"),
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Max, "TestColumn2"),
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Min, "TestColumn3"),
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Max, "TestColumn3"),
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Min, "TestColumn4"),
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Max, "TestColumn4"),
            };

            MethodInfo method = totalPanel.GetType().GetMethod("DoAggregation", BindingFlags.Instance | BindingFlags.NonPublic);
            method.Invoke(totalPanel, new object[] { enumerator, totalCells });

            Assert.AreEqual(6, totalCells[0].Result);
            Assert.AreEqual(62.1m, totalCells[1].Result);
            Assert.AreEqual("abcjkldef", totalCells[2].Result);
            Assert.AreEqual(4, totalCells[3].Result);
            Assert.AreEqual(4, totalCells[4].Result);
            Assert.AreEqual((double)6 / 4, totalCells[5].Result);
            Assert.AreEqual(62.1 / 4, totalCells[6].Result);
            Assert.AreEqual(1, totalCells[7].Result);
            Assert.AreEqual(3, totalCells[8].Result);
            Assert.AreEqual(10.5m, totalCells[9].Result);
            Assert.AreEqual(30.9m, totalCells[10].Result);
            Assert.AreEqual("abc", totalCells[11].Result);
            Assert.AreEqual("jkl", totalCells[12].Result);
            Assert.AreEqual(false, totalCells[13].Result);
            Assert.AreEqual(true, totalCells[14].Result);

            // Reset all results before next test
            foreach (ExcelTotalsPanel.ParsedAggregationFunc totalCell in totalCells)
            {
                totalCell.Result = null;
            }

            IList<Test> data = GetTestData();

            enumerator = EnumeratorFactory.Create(data);
            totalCells.Add(new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Sum, "Result.Amount"));

            method.Invoke(totalPanel, new object[] { enumerator, totalCells });

            Assert.AreEqual(6, totalCells[0].Result);
            Assert.AreEqual(62.1m, totalCells[1].Result);
            Assert.AreEqual("abcjkldef", totalCells[2].Result);
            Assert.AreEqual(4, totalCells[3].Result);
            Assert.AreEqual(4, totalCells[4].Result);
            Assert.AreEqual((double)6 / 4, totalCells[5].Result);
            Assert.AreEqual(62.1 / 4, totalCells[6].Result);
            Assert.AreEqual(1, totalCells[7].Result);
            Assert.AreEqual(3, totalCells[8].Result);
            Assert.AreEqual(10.5m, totalCells[9].Result);
            Assert.AreEqual(30.9m, totalCells[10].Result);
            Assert.AreEqual("abc", totalCells[11].Result);
            Assert.AreEqual("jkl", totalCells[12].Result);
            Assert.AreEqual(false, totalCells[13].Result);
            Assert.AreEqual(true, totalCells[14].Result);
            Assert.AreEqual(410.59m, totalCells[15].Result);
        }

        [TestMethod]
        public void TestDoAggregationWithEmptyData()
        {
            var data = new List<Test>();
            var totalPanel = new ExcelTotalsPanel(data, Substitute.For<IXLNamedRange>(), Substitute.For<object>(), Substitute.For<ITemplateProcessor>());
            IEnumerator enumerator = EnumeratorFactory.Create(data);
            IList<ExcelTotalsPanel.ParsedAggregationFunc> totalCells = new List<ExcelTotalsPanel.ParsedAggregationFunc>
            {
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Sum, "TestColumn1"),
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Count, "TestColumn1"),
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Avg, "TestColumn1"),
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Min, "TestColumn1"),
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Max, "TestColumn1"),
            };

            MethodInfo method = totalPanel.GetType().GetMethod("DoAggregation", BindingFlags.Instance | BindingFlags.NonPublic);
            method.Invoke(totalPanel, new object[] { enumerator, totalCells });

            Assert.AreEqual(0, totalCells[0].Result);
            Assert.AreEqual(0, totalCells[1].Result);
            Assert.AreEqual(0, totalCells[2].Result);
            Assert.IsNull(totalCells[3].Result);
            Assert.IsNull(totalCells[4].Result);
        }

        [TestMethod]
        public void TestDoAggregationWithBadData()
        {
            IList<Test> data = GetTestData();

            var totalPanel = new ExcelTotalsPanel(data, Substitute.For<IXLNamedRange>(), Substitute.For<object>(), Substitute.For<ITemplateProcessor>());
            IEnumerator enumerator = EnumeratorFactory.Create(data);
            IList<ExcelTotalsPanel.ParsedAggregationFunc> totalCells = new List<ExcelTotalsPanel.ParsedAggregationFunc>
            {
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Sum, "TestColumn4"),
            };

            MethodInfo method = totalPanel.GetType().GetMethod("DoAggregation", BindingFlags.Instance | BindingFlags.NonPublic);
            ExceptionAssert.ThrowsBaseException<RuntimeBinderException>(() => method.Invoke(totalPanel, new object[] { enumerator, totalCells }));

            enumerator.Reset();
            totalCells = new List<ExcelTotalsPanel.ParsedAggregationFunc>
            {
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Min, "BadColumn"),
            };
            ExceptionAssert.ThrowsBaseException<InvalidOperationException>(() => method.Invoke(totalPanel, new object[] { enumerator, totalCells }),
                "For Min and Max aggregation functions data items must implement IComparable interface");

            enumerator.Reset();
            totalCells = new List<ExcelTotalsPanel.ParsedAggregationFunc>
            {
                new ExcelTotalsPanel.ParsedAggregationFunc((AggregateFunction)6, "TestColumn1"),
            };
            ExceptionAssert.ThrowsBaseException<NotSupportedException>(() => method.Invoke(totalPanel, new object[] { enumerator, totalCells }),
                "Unsupportable aggregation function");
        }

        [TestMethod]
        public void TestDoAggregationWithComplexType()
        {
            var data = new List<Test>();
            var totalPanel = new ExcelTotalsPanel(data, Substitute.For<IXLNamedRange>(), Substitute.For<object>(), Substitute.For<ITemplateProcessor>());
            IEnumerator enumerator = EnumeratorFactory.Create(data);
            IList<ExcelTotalsPanel.ParsedAggregationFunc> totalCells = new List<ExcelTotalsPanel.ParsedAggregationFunc>
            {
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Sum, "TestColumn1"),
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Count, "TestColumn1"),
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Avg, "TestColumn1"),
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Min, "TestColumn1"),
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Max, "TestColumn1"),
            };
        }

        [TestMethod]
        public void TestCustomAggregation()
        {
            IList<Test> data = GetTestData();

            var totalPanel = new ExcelTotalsPanel(data, Substitute.For<IXLNamedRange>(), new TestReportForAggregation(), Substitute.For<ITemplateProcessor>());
            IEnumerator enumerator = EnumeratorFactory.Create(data);
            IList<ExcelTotalsPanel.ParsedAggregationFunc> totalCells = new List<ExcelTotalsPanel.ParsedAggregationFunc>
            {
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Custom, "TestColumn2")
                {
                    CustomFunc = "CustomAggregation",
                },
            };

            MethodInfo method = totalPanel.GetType().GetMethod("DoAggregation", BindingFlags.Instance | BindingFlags.NonPublic);
            method.Invoke(totalPanel, new object[] { enumerator, totalCells });
            Assert.AreEqual(24.18125m, totalCells.First().Result);

            enumerator.Reset();
            totalCells.First().CustomFunc = null;
            ExceptionAssert.ThrowsBaseException<InvalidOperationException>(() => method.Invoke(totalPanel, new object[] { enumerator, totalCells }),
                "The custom type of aggregation is specified in the template but custom function is missing");

            enumerator.Reset();
            totalCells.First().CustomFunc = string.Empty;
            ExceptionAssert.ThrowsBaseException<InvalidOperationException>(() => method.Invoke(totalPanel, new object[] { enumerator, totalCells }),
                "The custom type of aggregation is specified in the template but custom function is missing");

            enumerator.Reset();
            totalCells.First().CustomFunc = " ";
            ExceptionAssert.ThrowsBaseException<InvalidOperationException>(() => method.Invoke(totalPanel, new object[] { enumerator, totalCells }),
                "The custom type of aggregation is specified in the template but custom function is missing");

            enumerator.Reset();
            totalCells.First().CustomFunc = "BadMethod";
            ExceptionAssert.ThrowsBaseException<MethodNotFoundException>(() => method.Invoke(totalPanel, new object[] { enumerator, totalCells }),
                $"Cannot find public instance method \"BadMethod\" in type \"{nameof(TestReportForAggregation)}\"");
        }

        [TestMethod]
        public void TestAggregationPostOperation()
        {
            IList<Test> data = GetTestData();

            var totalPanel = new ExcelTotalsPanel(data, Substitute.For<IXLNamedRange>(), new TestReportForAggregation(), Substitute.For<ITemplateProcessor>());
            IEnumerator enumerator = EnumeratorFactory.Create(data);
            IList<ExcelTotalsPanel.ParsedAggregationFunc> totalCells = new List<ExcelTotalsPanel.ParsedAggregationFunc>
            {
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Sum, "TestColumn2") { PostProcessFunction = "PostSumOperation" },
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Min, "TestColumn3") { PostProcessFunction = "PostMinOperation" },
                new ExcelTotalsPanel.ParsedAggregationFunc(AggregateFunction.Custom, "TestColumn2")
                {
                    CustomFunc = "CustomAggregation",
                    PostProcessFunction = "PostCustomAggregation",
                },
            };

            MethodInfo method = totalPanel.GetType().GetMethod("DoAggregation", BindingFlags.Instance | BindingFlags.NonPublic);
            method.Invoke(totalPanel, new object[] { enumerator, totalCells });
            Assert.AreEqual(22.033.ToString("F3"), totalCells[0].Result);
            Assert.AreEqual("ABC", totalCells[1].Result);
            Assert.AreEqual(24, totalCells[2].Result);
        }

        [TestMethod]
        public void TestParseTotalCells()
        {
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLRange range = ws.Range(1, 1, 1, 6);
            range.AddToNamed("Test", XLScope.Worksheet);

            ws.Cell(1, 1).Value = "Plain text";
            ws.Cell(1, 2).Value = "{Sum(di:Amount)}";
            ws.Cell(1, 3).Value = "{ Custom(DI:Amount, CustomFunc)  }";
            ws.Cell(1, 4).Value = "{Min(di:Value, CustomFunc, PostFunc)}";
            ws.Cell(1, 5).Value = "Text {count(Number)} {Text} {AVG( di:Value, ,  PostFunc )} Text {Max(Val)}";
            ws.Cell(1, 6).Value = "{Mix(di:Amount)}";
            ws.Cell(1, 7).Value = "{Sum(di:Amount)}";

            var templateProcessor = Substitute.For<ITemplateProcessor>();
            templateProcessor.LeftTemplateBorder.Returns("{");
            templateProcessor.RightTemplateBorder.Returns("}");
            templateProcessor.MemberLabelSeparator.Returns(":");
            templateProcessor.DataItemMemberLabel.Returns("di");

            var report = new TestReport
            {
                TemplateProcessor = templateProcessor,
                Workbook = wb
            };

            var panel = new ExcelTotalsPanel("Stub", ws.NamedRange("Test"), report, report.TemplateProcessor);
            MethodInfo method = panel.GetType().GetMethod("ParseTotalCells", BindingFlags.Instance | BindingFlags.NonPublic);
            var result = (IDictionary<IXLCell, IList<ExcelTotalsPanel.ParsedAggregationFunc>>)method.Invoke(panel, null);

            Assert.AreEqual(4, result.Count);
            Assert.AreEqual("Plain text", ws.Cell(1, 1).Value);
            Assert.AreEqual("{Mix(di:Amount)}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{Sum(di:Amount)}", ws.Cell(1, 7).Value);

            Assert.AreEqual("{0}", ws.Cell(1, 2).Value);
            Assert.AreEqual(1, result[ws.Cell(1, 2)].Count);
            Assert.AreEqual(AggregateFunction.Sum, result[ws.Cell(1, 2)].First().AggregateFunction);
            Assert.AreEqual("Amount", result[ws.Cell(1, 2)].First().ColumnName);
            Assert.IsNull(result[ws.Cell(1, 2)].First().CustomFunc);
            Assert.IsNull(result[ws.Cell(1, 2)].First().PostProcessFunction);
            Assert.IsNull(result[ws.Cell(1, 2)].First().Result);

            Assert.AreEqual("{0}", ws.Cell(1, 3).Value);
            Assert.AreEqual(1, result[ws.Cell(1, 3)].Count);
            Assert.AreEqual(AggregateFunction.Custom, result[ws.Cell(1, 3)].First().AggregateFunction);
            Assert.AreEqual("Amount", result[ws.Cell(1, 3)].First().ColumnName);
            Assert.AreEqual("CustomFunc", result[ws.Cell(1, 3)].First().CustomFunc);
            Assert.IsNull(result[ws.Cell(1, 3)].First().PostProcessFunction);
            Assert.IsNull(result[ws.Cell(1, 3)].First().Result);

            Assert.AreEqual("{0}", ws.Cell(1, 4).Value);
            Assert.AreEqual(1, result[ws.Cell(1, 4)].Count);
            Assert.AreEqual(AggregateFunction.Min, result[ws.Cell(1, 4)].First().AggregateFunction);
            Assert.AreEqual("Value", result[ws.Cell(1, 4)].First().ColumnName);
            Assert.AreEqual("CustomFunc", result[ws.Cell(1, 4)].First().CustomFunc);
            Assert.AreEqual("PostFunc", result[ws.Cell(1, 4)].First().PostProcessFunction);
            Assert.IsNull(result[ws.Cell(1, 4)].First().Result);

            Assert.AreEqual("Text {0} {Text} {1} Text {2}", ws.Cell(1, 5).Value);
            Assert.AreEqual(3, result[ws.Cell(1, 5)].Count);
            Assert.AreEqual(AggregateFunction.Count, result[ws.Cell(1, 5)][0].AggregateFunction);
            Assert.AreEqual("Number", result[ws.Cell(1, 5)][0].ColumnName);
            Assert.IsNull(result[ws.Cell(1, 5)][0].CustomFunc);
            Assert.IsNull(result[ws.Cell(1, 5)][0].PostProcessFunction);
            Assert.IsNull(result[ws.Cell(1, 5)][0].Result);
            Assert.AreEqual(AggregateFunction.Avg, result[ws.Cell(1, 5)][1].AggregateFunction);
            Assert.AreEqual("Value", result[ws.Cell(1, 5)][1].ColumnName);
            Assert.IsNull(result[ws.Cell(1, 5)][1].CustomFunc);
            Assert.AreEqual("PostFunc", result[ws.Cell(1, 5)][1].PostProcessFunction);
            Assert.IsNull(result[ws.Cell(1, 5)][1].Result);
            Assert.AreEqual(AggregateFunction.Max, result[ws.Cell(1, 5)][2].AggregateFunction);
            Assert.AreEqual("Val", result[ws.Cell(1, 5)][2].ColumnName);
            Assert.IsNull(result[ws.Cell(1, 5)][2].CustomFunc);
            Assert.IsNull(result[ws.Cell(1, 5)][2].PostProcessFunction);
            Assert.IsNull(result[ws.Cell(1, 5)][2].Result);
        }

        [TestMethod]
        public void TestParseTotalCellsErrors()
        {
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLRange range = ws.Range(1, 1, 1, 1);
            range.AddToNamed("Test", XLScope.Worksheet);

            ws.Cell(1, 1).Value = "<Sum( )>";

            var templateProcessor = Substitute.For<ITemplateProcessor>();
            templateProcessor.LeftTemplateBorder.Returns("<");
            templateProcessor.RightTemplateBorder.Returns(">");
            templateProcessor.MemberLabelSeparator.Returns("-");
            templateProcessor.DataItemMemberLabel.Returns("d");

            var report = new TestReport
            {
                TemplateProcessor = templateProcessor,
                Workbook = wb
            };

            var panel = new ExcelTotalsPanel("Stub", ws.NamedRange("Test"), report, report.TemplateProcessor);
            MethodInfo method = panel.GetType().GetMethod("ParseTotalCells", BindingFlags.Instance | BindingFlags.NonPublic);

            ExceptionAssert.ThrowsBaseException<InvalidOperationException>(() => method.Invoke(panel, null), "\"ColumnName\" parameter in aggregation function cannot be empty");

            ws.Cell(1, 1).Value = "<Sum(di-Val, fn1, fn2, fn3)>";
            ExceptionAssert.ThrowsBaseException<InvalidOperationException>(() => method.Invoke(panel, null), "Aggregation function must have at least one but no more than 3 parameters");

            ws.Cell(1, 1).Value = "<Sum( , fn1, fn2)>";
            ExceptionAssert.ThrowsBaseException<InvalidOperationException>(() => method.Invoke(panel, null), "\"ColumnName\" parameter in aggregation function cannot be empty");
        }

        private IList<Test> GetTestData()
        {
            return new List<Test>
            {
                new Test(3, 20.7m, "abc", false) { Result = new ComplexType { Amount = 155.05m }},
                new Test(1, 10.5m, "jkl", true) { Result = new ComplexType() },
                new Test(null, null, null, null) { Result = new ComplexType() },
                new Test(2, 30.9m, "def", false) { Result = new ComplexType { Amount = 255.54m }},
            };
        }

        private class Test
        {
            public Test(int? testColumn1, decimal? testColumn2, string testColumn3, bool? testColumn4)
            {
                TestColumn1 = testColumn1;
                TestColumn2 = testColumn2;
                TestColumn3 = testColumn3;
                TestColumn4 = testColumn4;
                BadColumn = new Test();
            }

            private Test()
            {
            }

            public int? TestColumn1 { get; }
            public decimal? TestColumn2 { get; }
            public string TestColumn3 { get; }
            public bool? TestColumn4 { get; }
            public Test BadColumn { get; }
            public ComplexType Result { get; set; }
        }

        private class ComplexType
        {
            public decimal Amount { get; set; }
        }

        private class TestReportForAggregation
        {
            public decimal CustomAggregation(decimal result, decimal currentValue, int itemNumber)
            {
                return (result + currentValue) / 2 + itemNumber;
            }

            public string PostSumOperation(decimal result, int itemsCount)
            {
                return ((result + itemsCount) / 3).ToString("F3");
            }

            public string PostMinOperation(string result, int itemsCount)
            {
                return result.ToUpper();
            }

            public int PostCustomAggregation(decimal result, int itemsCount)
            {
                return (int)decimal.Round(result, 0);
            }
        }
    }
}