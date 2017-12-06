using ClosedXML.Excel;
using ExcelReporter.Enumerators;
using ExcelReporter.Enums;
using ExcelReporter.Rendering.Panels.ExcelPanels;
using ExcelReporter.Rendering.TemplateProcessors;
using ExcelReporter.Reports;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using ExcelReporter.Exceptions;
using DataTable = System.Data.DataTable;

namespace ExcelReporter.Tests.Rendering.Panels.ExcelPanels
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

            var totalPanel = new ExcelTotalsPanel(dataTable, Substitute.For<IXLNamedRange>(), Substitute.For<IExcelReport>());
            IEnumerator enumerator = EnumeratorFactory.Create(dataTable);
            IList<ExcelTotalsPanel.TotalCellInfo> totalCells = new List<ExcelTotalsPanel.TotalCellInfo>
            {
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Sum, "TestColumn1"),
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Sum, "TestColumn2"),
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Sum, "TestColumn3"),
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Count, "TestColumn1"),
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Count, "TestColumn3"),
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Avg, "TestColumn1"),
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Avg, "TestColumn2"),
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Min, "TestColumn1"),
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Max, "TestColumn1"),
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Min, "TestColumn2"),
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Max, "TestColumn2"),
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Min, "TestColumn3"),
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Max, "TestColumn3"),
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Min, "TestColumn4"),
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Max, "TestColumn4"),
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
            foreach (ExcelTotalsPanel.TotalCellInfo totalCell in totalCells)
            {
                totalCell.Result = null;
            }

            IList<Test> data = GetTestData();

            enumerator = EnumeratorFactory.Create(data);
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
        }

        [TestMethod]
        public void TestDoAggregationWithEmptyData()
        {
            var data = new List<Test>();
            var totalPanel = new ExcelTotalsPanel(data, Substitute.For<IXLNamedRange>(), Substitute.For<IExcelReport>());
            IEnumerator enumerator = EnumeratorFactory.Create(data);
            IList<ExcelTotalsPanel.TotalCellInfo> totalCells = new List<ExcelTotalsPanel.TotalCellInfo>
            {
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Sum, "TestColumn1"),
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Count, "TestColumn1"),
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Avg, "TestColumn1"),
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Min, "TestColumn1"),
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Max, "TestColumn1"),
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

            var totalPanel = new ExcelTotalsPanel(data, Substitute.For<IXLNamedRange>(), Substitute.For<IExcelReport>());
            IEnumerator enumerator = EnumeratorFactory.Create(data);
            IList<ExcelTotalsPanel.TotalCellInfo> totalCells = new List<ExcelTotalsPanel.TotalCellInfo>
            {
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Sum, "TestColumn4"),
            };

            MethodInfo method = totalPanel.GetType().GetMethod("DoAggregation", BindingFlags.Instance | BindingFlags.NonPublic);
            ExceptionAssert.ThrowsBaseException<RuntimeBinderException>(() => method.Invoke(totalPanel, new object[] { enumerator, totalCells }));

            enumerator.Reset();
            totalCells = new List<ExcelTotalsPanel.TotalCellInfo>
            {
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Min, "BadColumn"),
            };
            ExceptionAssert.ThrowsBaseException<InvalidOperationException>(() => method.Invoke(totalPanel, new object[] { enumerator, totalCells }),
                "For Min and Max aggregation functions data items must implement IComparable interface");

            enumerator.Reset();
            totalCells = new List<ExcelTotalsPanel.TotalCellInfo>
            {
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), (AggregateFunction)6, "TestColumn1"),
            };
            ExceptionAssert.ThrowsBaseException<NotSupportedException>(() => method.Invoke(totalPanel, new object[] { enumerator, totalCells }),
                "Unsupportable aggregation function");
        }

        [TestMethod]
        public void TestCustomAggregation()
        {
            IList<Test> data = GetTestData();

            var totalPanel = new ExcelTotalsPanel(data, Substitute.For<IXLNamedRange>(), new TestReportForAggregation());
            IEnumerator enumerator = EnumeratorFactory.Create(data);
            IList<ExcelTotalsPanel.TotalCellInfo> totalCells = new List<ExcelTotalsPanel.TotalCellInfo>
            {
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Custom, "TestColumn2")
                {
                    CustomFunc = "CustomAggregation",
                },
            };

            MethodInfo method = totalPanel.GetType().GetMethod("DoAggregation", BindingFlags.Instance | BindingFlags.NonPublic);
            method.Invoke(totalPanel, new object[] {enumerator, totalCells});
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
        public void TestPostOperation()
        {
            IList<Test> data = GetTestData();

            var totalPanel = new ExcelTotalsPanel(data, Substitute.For<IXLNamedRange>(), new TestReportForAggregation());
            IEnumerator enumerator = EnumeratorFactory.Create(data);
            IList<ExcelTotalsPanel.TotalCellInfo> totalCells = new List<ExcelTotalsPanel.TotalCellInfo>
            {
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Sum, "TestColumn2") { PostProcessFunction = "PostSumOperation" },
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Min, "TestColumn3") { PostProcessFunction = "PostMinOperation" },
                new ExcelTotalsPanel.TotalCellInfo(Substitute.For<IXLCell>(), AggregateFunction.Custom, "TestColumn2")
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

        private IList<Test> GetTestData()
        {
            return new List<Test>
            {
                new Test(3, 20.7m, "abc", false),
                new Test(1, 10.5m, "jkl", true),
                new Test(null, null, null, null),
                new Test(2, 30.9m, "def", false),
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
        }

        private class TestReportForAggregation : IExcelReport
        {
            public ITemplateProcessor TemplateProcessor { get; set; }
            public XLWorkbook Workbook { get; set; }

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

            public void Run()
            {
                throw new NotImplementedException();
            }
        }
    }
}