using ClosedXML.Excel;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using ExcelReportGenerator.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;
using System;
using System.Reflection;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels
{
    [TestClass]
    public class ExcelDataSourcePanelTest
    {
        [TestMethod]
        public void TestGroupResultVertical()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 9, 6);

            ws.Cell(2, 2).Value = "One";
            ws.Cell(3, 2).Value = "One";
            ws.Cell(4, 2).Value = "Two";
            ws.Cell(5, 2).Value = "Three";
            ws.Cell(6, 2).Value = "Three";
            ws.Cell(7, 2).Value = "Three";
            ws.Cell(8, 2).Value = "Four";
            ws.Cell(9, 2).Value = "Five";

            ws.Range(5, 2, 6, 2).Merge();

            ws.Cell(2, 3).Value = "Orange";
            ws.Cell(3, 3).Value = "Apple";
            ws.Cell(4, 3).Value = "Apple";
            ws.Cell(5, 3).Value = string.Empty;
            ws.Cell(6, 3).Value = null;
            ws.Cell(8, 3).Value = "Pear";
            ws.Cell(9, 3).Value = "Pear";

            ws.Cell(2, 4).Value = true;
            ws.Cell(3, 4).Value = true;
            ws.Cell(4, 4).Value = 1;
            ws.Cell(5, 4).Value = null;
            ws.Cell(7, 4).Value = 0;
            ws.Cell(8, 4).Value = false;
            ws.Cell(9, 4).Value = false;

            ws.Cell(2, 5).Value = 1;
            ws.Cell(3, 5).Value = 1;
            ws.Cell(4, 5).Value = 1;
            ws.Cell(5, 5).Value = 56;
            ws.Cell(6, 5).Value = 56.1;
            ws.Cell(7, 5).Value = 56;
            ws.Cell(8, 5).Value = 77.7;
            ws.Cell(9, 5).Value = 77.7m;

            ws.Range(3, 5, 4, 5).Merge();

            ws.Cell(2, 6).Value = new DateTime(2018, 2, 18);
            ws.Cell(3, 6).Value = new DateTime(2018, 2, 20);
            ws.Cell(4, 6).Value = new DateTime(2018, 2, 20);
            ws.Cell(5, 6).Value = new DateTime(2018, 2, 18);
            ws.Cell(6, 6).Value = null;
            ws.Cell(8, 6).Value = new DateTime(2018, 2, 21);
            ws.Cell(9, 6).Value = new DateTime(2018, 2, 21).ToString();

            var panel = new ExcelDataSourcePanel("Stub", Substitute.For<IXLNamedRange>(), new object(), Substitute.For<ITemplateProcessor>())
            {
                GroupBy = "1,2, 3 , 4,5"
            };

            var method = panel.GetType().GetMethod("GroupResult", BindingFlags.Instance | BindingFlags.NonPublic);
            method.Invoke(panel, new[] { range });

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(ExcelDataSourcePanelTest), nameof(TestGroupResultVertical)), wb);

            //wb.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestGroupResultHorizontal()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 6, 9);

            ws.Cell(2, 2).Value = "One";
            ws.Cell(2, 3).Value = "One";
            ws.Cell(2, 4).Value = "Two";
            ws.Cell(2, 5).Value = "Three";
            ws.Cell(2, 6).Value = "Three";
            ws.Cell(2, 7).Value = "Three";
            ws.Cell(2, 8).Value = "Four";
            ws.Cell(2, 9).Value = "Five";

            ws.Range(2, 5, 2, 6).Merge();

            ws.Cell(3, 2).Value = "Orange";
            ws.Cell(3, 3).Value = "Apple";
            ws.Cell(3, 4).Value = "Apple";
            ws.Cell(3, 5).Value = string.Empty;
            ws.Cell(3, 6).Value = null;
            ws.Cell(3, 8).Value = "Pear";
            ws.Cell(3, 9).Value = "Pear";

            ws.Cell(4, 2).Value = true;
            ws.Cell(4, 3).Value = true;
            ws.Cell(4, 4).Value = 1;
            ws.Cell(4, 5).Value = null;
            ws.Cell(4, 7).Value = 0;
            ws.Cell(4, 8).Value = false;
            ws.Cell(4, 9).Value = false;

            ws.Cell(5, 2).Value = 1;
            ws.Cell(5, 3).Value = 1;
            ws.Cell(5, 4).Value = 1;
            ws.Cell(5, 5).Value = 56;
            ws.Cell(5, 6).Value = 56.1;
            ws.Cell(5, 7).Value = 56;
            ws.Cell(5, 8).Value = 77.7;
            ws.Cell(5, 9).Value = 77.7m;

            ws.Range(5, 3, 5, 4).Merge();

            ws.Cell(6, 2).Value = new DateTime(2018, 2, 18);
            ws.Cell(6, 3).Value = new DateTime(2018, 2, 20);
            ws.Cell(6, 4).Value = new DateTime(2018, 2, 20);
            ws.Cell(6, 5).Value = new DateTime(2018, 2, 18);
            ws.Cell(6, 6).Value = null;
            ws.Cell(6, 8).Value = new DateTime(2018, 2, 21);
            ws.Cell(6, 9).Value = new DateTime(2018, 2, 21).ToString();

            var panel = new ExcelDataSourcePanel("Stub", Substitute.For<IXLNamedRange>(), new object(), Substitute.For<ITemplateProcessor>())
            {
                GroupBy = "1,2, 3 , 4,5",
                Type = PanelType.Horizontal,
            };

            var method = panel.GetType().GetMethod("GroupResult", BindingFlags.Instance | BindingFlags.NonPublic);
            method.Invoke(panel, new[] { range });

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(ExcelDataSourcePanelTest), nameof(TestGroupResultHorizontal)), wb);

            //wb.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestIfGroupByPropertyIsEmpty()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 6, 3);

            ws.Cell(2, 2).Value = "One";
            ws.Cell(3, 2).Value = "One";
            ws.Cell(4, 2).Value = "Two";
            ws.Cell(5, 2).Value = "Three";
            ws.Cell(6, 2).Value = "Three";
            ws.Cell(7, 2).Value = "Three";
            ws.Cell(8, 2).Value = "Four";
            ws.Cell(9, 2).Value = "Five";

            ws.Cell(2, 3).Value = "Orange";
            ws.Cell(3, 3).Value = "Apple";
            ws.Cell(4, 3).Value = "Apple";
            ws.Cell(5, 3).Value = string.Empty;
            ws.Cell(6, 3).Value = null;
            ws.Cell(8, 3).Value = "Pear";
            ws.Cell(9, 3).Value = "Pear";

            var panel = new ExcelDataSourcePanel("Stub", Substitute.For<IXLNamedRange>(), new object(), Substitute.For<ITemplateProcessor>()) { GroupBy = null };

            var method = panel.GetType().GetMethod("GroupResult", BindingFlags.Instance | BindingFlags.NonPublic);
            method.Invoke(panel, new[] { range });

            Assert.AreEqual(0, ws.MergedRanges.Count);

            panel = new ExcelDataSourcePanel("Stub", Substitute.For<IXLNamedRange>(), new object(), Substitute.For<ITemplateProcessor>()) { GroupBy = string.Empty };
            method.Invoke(panel, new[] { range });

            Assert.AreEqual(0, ws.MergedRanges.Count);

            panel = new ExcelDataSourcePanel("Stub", Substitute.For<IXLNamedRange>(), new object(), Substitute.For<ITemplateProcessor>()) { GroupBy = "  " };
            method.Invoke(panel, new[] { range });

            Assert.AreEqual(0, ws.MergedRanges.Count);

            //wb.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestIfGroupByPropertyIsInvalid()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 2);

            ws.Cell(2, 2).Value = "One";
            ws.Cell(3, 2).Value = "One";

            var panel = new ExcelDataSourcePanel("Stub", Substitute.For<IXLNamedRange>(), new object(), Substitute.For<ITemplateProcessor>()) { GroupBy = "str" };
            var method = panel.GetType().GetMethod("GroupResult", BindingFlags.Instance | BindingFlags.NonPublic);

            ExceptionAssert.ThrowsBaseException<InvalidCastException>(() => method.Invoke(panel, new[] { range }), $"Parse \"GroupBy\" property failed. Cannot convert value \"str\" to {nameof(Int32)}");

            panel = new ExcelDataSourcePanel("Stub", Substitute.For<IXLNamedRange>(), new object(), Substitute.For<ITemplateProcessor>()) { GroupBy = "1, 1.4" };

            ExceptionAssert.ThrowsBaseException<InvalidCastException>(() => method.Invoke(panel, new[] { range }), $"Parse \"GroupBy\" property failed. Cannot convert value \"1.4\" to {nameof(Int32)}");
        }
    }
}