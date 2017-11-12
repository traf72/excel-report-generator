using System;
using System.Linq;
using ClosedXML.Excel;
using ExcelReporter.Enums;
using ExcelReporter.Implementations.Panels.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Implementations.Panels.Excel.PanelRenderTests.DataSourcePanelRenderTests
{
    [TestClass]
    public class DataSourcePanelIEnumerableRenderTest
    {
        [TestMethod]
        public void TestRenderIEnumerableVerticalCellsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

            ws.Cell(2, 4).DataType = XLCellValues.Number;

            ws.Cell(2, 2).Value = "{m:Concat(di:Name, m:Format(di:Date, dd.MM.yyyy))}";
            ws.Cell(2, 3).Value = "{di:Date}";
            ws.Cell(2, 4).Value = "{m:Multiply(di:Sum, 5)}";
            ws.Cell(2, 5).Value = "{di:Contacts}";
            ws.Cell(3, 2).Value = "{di:Contacts.Phone}";
            ws.Cell(3, 3).Value = "{di:Contacts.Fax}";
            ws.Cell(3, 4).Value = "{p:StrParam}";

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(4, 1).Value = "{di:Name}";
            ws.Cell(1, 6).Value = "{di:Name}";
            ws.Cell(4, 6).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 6).Value = "{di:Name}";
            ws.Cell(1, 4).Value = "{di:Name}";
            ws.Cell(4, 4).Value = "{di:Name}";

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("TestRange"), report);
            panel.Render();

            Assert.AreEqual(29, ws.CellsUsed().Count());
            Assert.AreEqual("Test1_01.11.2017", ws.Cell(2, 2).Value);
            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(2, 3).Value);
            Assert.AreEqual(278.8, ws.Cell(2, 4).Value);
            Assert.AreEqual("15_345", ws.Cell(2, 5).Value);
            Assert.AreEqual(15d, ws.Cell(3, 2).Value);
            Assert.AreEqual(345d, ws.Cell(3, 3).Value);
            Assert.AreEqual("String parameter", ws.Cell(3, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLCellValues.Number, ws.Cell(2, 4).DataType);

            Assert.AreEqual("Test2_02.11.2017", ws.Cell(4, 2).Value);
            Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(4, 3).Value);
            Assert.AreEqual(550d, ws.Cell(4, 4).Value);
            Assert.AreEqual("76_753465", ws.Cell(4, 5).Value);
            Assert.AreEqual(76d, ws.Cell(5, 2).Value);
            Assert.AreEqual(753465d, ws.Cell(5, 3).Value);
            Assert.AreEqual("String parameter", ws.Cell(5, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLCellValues.Number, ws.Cell(4, 4).DataType);

            Assert.AreEqual("Test3_03.11.2017", ws.Cell(6, 2).Value);
            Assert.AreEqual(new DateTime(2017, 11, 3), ws.Cell(6, 3).Value);
            Assert.AreEqual(27504d, ws.Cell(6, 4).Value);
            Assert.AreEqual("1533_5456", ws.Cell(6, 5).Value);
            Assert.AreEqual(1533d, ws.Cell(7, 2).Value);
            Assert.AreEqual(5456d, ws.Cell(7, 3).Value);
            Assert.AreEqual("String parameter", ws.Cell(7, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLCellValues.Number, ws.Cell(6, 4).DataType);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(8, 4).Value);

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderIEnumerableVerticalRowsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

            ws.Cell(2, 4).DataType = XLCellValues.Number;

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";
            ws.Cell(2, 4).Value = "{di:Sum}";
            ws.Cell(2, 5).Value = "{di:Contacts}";
            ws.Cell(3, 2).Value = "{di:Contacts.Phone}";
            ws.Cell(3, 3).Value = "{di:Contacts.Fax}";
            ws.Cell(3, 4).Value = "{p:StrParam}";

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(4, 1).Value = "{di:Name}";
            ws.Cell(1, 6).Value = "{di:Name}";
            ws.Cell(4, 6).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 6).Value = "{di:Name}";
            ws.Cell(1, 4).Value = "{di:Name}";
            ws.Cell(4, 4).Value = "{di:Name}";

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("TestRange"), report)
            {
                ShiftType = ShiftType.Row,
            };
            panel.Render();

            Assert.AreEqual(29, ws.CellsUsed().Count());
            Assert.AreEqual("Test1", ws.Cell(2, 2).Value);
            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(2, 3).Value);
            Assert.AreEqual(55.76, ws.Cell(2, 4).Value);
            Assert.AreEqual("15_345", ws.Cell(2, 5).Value);
            Assert.AreEqual(15d, ws.Cell(3, 2).Value);
            Assert.AreEqual(345d, ws.Cell(3, 3).Value);
            Assert.AreEqual("String parameter", ws.Cell(3, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLCellValues.Number, ws.Cell(2, 4).DataType);

            Assert.AreEqual("Test2", ws.Cell(4, 2).Value);
            Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(4, 3).Value);
            Assert.AreEqual(110d, ws.Cell(4, 4).Value);
            Assert.AreEqual("76_753465", ws.Cell(4, 5).Value);
            Assert.AreEqual(76d, ws.Cell(5, 2).Value);
            Assert.AreEqual(753465d, ws.Cell(5, 3).Value);
            Assert.AreEqual("String parameter", ws.Cell(5, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLCellValues.Number, ws.Cell(4, 4).DataType);

            Assert.AreEqual("Test3", ws.Cell(6, 2).Value);
            Assert.AreEqual(new DateTime(2017, 11, 3), ws.Cell(6, 3).Value);
            Assert.AreEqual(5500.80, ws.Cell(6, 4).Value);
            Assert.AreEqual("1533_5456", ws.Cell(6, 5).Value);
            Assert.AreEqual(1533d, ws.Cell(7, 2).Value);
            Assert.AreEqual(5456d, ws.Cell(7, 3).Value);
            Assert.AreEqual("String parameter", ws.Cell(7, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLCellValues.Number, ws.Cell(6, 4).DataType);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(8, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(8, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(7, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(7, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(8, 4).Value);

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderIEnumerableVerticalNoShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

            ws.Cell(3, 5).Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);

            ws.Cell(2, 4).DataType = XLCellValues.Number;

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";
            ws.Cell(2, 4).Value = "{di:Sum}";
            ws.Cell(2, 5).Value = "{di:Contacts}";
            ws.Cell(3, 2).Value = "{di:Contacts.Phone}";
            ws.Cell(3, 3).Value = "{di:Contacts.Fax}";
            ws.Cell(3, 4).Value = "{p:StrParam}";

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(4, 1).Value = "{di:Name}";
            ws.Cell(1, 6).Value = "{di:Name}";
            ws.Cell(4, 6).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 6).Value = "{di:Name}";
            ws.Cell(1, 4).Value = "{di:Name}";
            ws.Cell(4, 4).Value = "{di:Name}";
            ws.Cell(8, 5).Value = "{di:Date}";

            ws.Cell(8, 5).Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            ws.Cell(8, 5).Style.Border.SetRightBorder(XLBorderStyleValues.Thin);
            ws.Cell(8, 5).Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
            ws.Cell(8, 5).Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

            ws.Cell(8, 5).Style.Border.SetTopBorderColor(XLColor.Red);
            ws.Cell(8, 5).Style.Border.SetRightBorderColor(XLColor.Red);
            ws.Cell(8, 5).Style.Border.SetBottomBorderColor(XLColor.Red);
            ws.Cell(8, 5).Style.Border.SetLeftBorderColor(XLColor.Red);

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("TestRange"), report)
            {
                ShiftType = ShiftType.NoShift,
            };
            panel.Render();

            Assert.AreEqual(29, ws.CellsUsed().Count());
            Assert.AreEqual("Test1", ws.Cell(2, 2).Value);
            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(2, 3).Value);
            Assert.AreEqual(55.76, ws.Cell(2, 4).Value);
            Assert.AreEqual("15_345", ws.Cell(2, 5).Value);
            Assert.AreEqual(15d, ws.Cell(3, 2).Value);
            Assert.AreEqual(345d, ws.Cell(3, 3).Value);
            Assert.AreEqual("String parameter", ws.Cell(3, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLCellValues.Number, ws.Cell(2, 4).DataType);

            Assert.AreEqual("Test2", ws.Cell(4, 2).Value);
            Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(4, 3).Value);
            Assert.AreEqual(110d, ws.Cell(4, 4).Value);
            Assert.AreEqual("76_753465", ws.Cell(4, 5).Value);
            Assert.AreEqual(76d, ws.Cell(5, 2).Value);
            Assert.AreEqual(753465d, ws.Cell(5, 3).Value);
            Assert.AreEqual("String parameter", ws.Cell(5, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLCellValues.Number, ws.Cell(4, 4).DataType);

            Assert.AreEqual("Test3", ws.Cell(6, 2).Value);
            Assert.AreEqual(new DateTime(2017, 11, 3), ws.Cell(6, 3).Value);
            Assert.AreEqual(5500.80, ws.Cell(6, 4).Value);
            Assert.AreEqual("1533_5456", ws.Cell(6, 5).Value);
            Assert.AreEqual(1533d, ws.Cell(7, 2).Value);
            Assert.AreEqual(5456d, ws.Cell(7, 3).Value);
            Assert.AreEqual("String parameter", ws.Cell(7, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLCellValues.Number, ws.Cell(6, 4).DataType);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            Assert.AreEqual("{di:Date}", ws.Cell(8, 5).Value);

            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 5).Style.Border.LeftBorder);

            Assert.AreEqual(XLColor.Black, ws.Cell(7, 5).Style.Border.BottomBorderColor);

            Assert.AreEqual(XLColor.Red, ws.Cell(8, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 5).Style.Border.LeftBorderColor);

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderIEnumerableHorizontalCellsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

            ws.Cell(2, 4).DataType = XLCellValues.Number;

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";
            ws.Cell(2, 4).Value = "{di:Sum}";
            ws.Cell(2, 5).Value = "{di:Contacts}";
            ws.Cell(3, 2).Value = "{di:Contacts.Phone}";
            ws.Cell(3, 3).Value = "{di:Contacts.Fax}";
            ws.Cell(3, 4).Value = "{p:StrParam}";

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(4, 1).Value = "{di:Name}";
            ws.Cell(1, 6).Value = "{di:Name}";
            ws.Cell(4, 6).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 6).Value = "{di:Name}";
            ws.Cell(1, 4).Value = "{di:Name}";
            ws.Cell(4, 4).Value = "{di:Name}";

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("TestRange"), report)
            {
                Type = PanelType.Horizontal,
            };
            panel.Render();

            Assert.AreEqual(29, ws.CellsUsed().Count());
            Assert.AreEqual("Test1", ws.Cell(2, 2).Value);
            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(2, 3).Value);
            Assert.AreEqual(55.76, ws.Cell(2, 4).Value);
            Assert.AreEqual("15_345", ws.Cell(2, 5).Value);
            Assert.AreEqual(15d, ws.Cell(3, 2).Value);
            Assert.AreEqual(345d, ws.Cell(3, 3).Value);
            Assert.AreEqual("String parameter", ws.Cell(3, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLCellValues.Number, ws.Cell(2, 4).DataType);

            Assert.AreEqual("Test2", ws.Cell(2, 6).Value);
            Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(2, 7).Value);
            Assert.AreEqual(110d, ws.Cell(2, 8).Value);
            Assert.AreEqual("76_753465", ws.Cell(2, 9).Value);
            Assert.AreEqual(76d, ws.Cell(3, 6).Value);
            Assert.AreEqual(753465d, ws.Cell(3, 7).Value);
            Assert.AreEqual("String parameter", ws.Cell(3, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.LeftBorder);
            Assert.AreEqual(XLCellValues.Number, ws.Cell(2, 8).DataType);

            Assert.AreEqual("Test3", ws.Cell(2, 10).Value);
            Assert.AreEqual(new DateTime(2017, 11, 3), ws.Cell(2, 11).Value);
            Assert.AreEqual(5500.80, ws.Cell(2, 12).Value);
            Assert.AreEqual("1533_5456", ws.Cell(2, 13).Value);
            Assert.AreEqual(1533d, ws.Cell(3, 10).Value);
            Assert.AreEqual(5456d, ws.Cell(3, 11).Value);
            Assert.AreEqual("String parameter", ws.Cell(3, 12).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 10).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 10).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 10).Style.Border.LeftBorder);
            Assert.AreEqual(XLCellValues.Number, ws.Cell(2, 12).DataType);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 14).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 4).Value);

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderIEnumerableHorizontalRowsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

            ws.Cell(2, 4).DataType = XLCellValues.Number;

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";
            ws.Cell(2, 4).Value = "{di:Sum}";
            ws.Cell(2, 5).Value = "{di:Contacts}";
            ws.Cell(3, 2).Value = "{di:Contacts.Phone}";
            ws.Cell(3, 3).Value = "{di:Contacts.Fax}";
            ws.Cell(3, 4).Value = "{p:StrParam}";

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(4, 1).Value = "{di:Name}";
            ws.Cell(1, 6).Value = "{di:Name}";
            ws.Cell(4, 6).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 6).Value = "{di:Name}";
            ws.Cell(1, 4).Value = "{di:Name}";
            ws.Cell(4, 4).Value = "{di:Name}";

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("TestRange"), report)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.Row,
            };
            panel.Render();

            Assert.AreEqual(29, ws.CellsUsed().Count());
            Assert.AreEqual("Test1", ws.Cell(2, 2).Value);
            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(2, 3).Value);
            Assert.AreEqual(55.76, ws.Cell(2, 4).Value);
            Assert.AreEqual("15_345", ws.Cell(2, 5).Value);
            Assert.AreEqual(15d, ws.Cell(3, 2).Value);
            Assert.AreEqual(345d, ws.Cell(3, 3).Value);
            Assert.AreEqual("String parameter", ws.Cell(3, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLCellValues.Number, ws.Cell(2, 4).DataType);

            Assert.AreEqual("Test2", ws.Cell(2, 6).Value);
            Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(2, 7).Value);
            Assert.AreEqual(110d, ws.Cell(2, 8).Value);
            Assert.AreEqual("76_753465", ws.Cell(2, 9).Value);
            Assert.AreEqual(76d, ws.Cell(3, 6).Value);
            Assert.AreEqual(753465d, ws.Cell(3, 7).Value);
            Assert.AreEqual("String parameter", ws.Cell(3, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.LeftBorder);
            Assert.AreEqual(XLCellValues.Number, ws.Cell(2, 8).DataType);

            Assert.AreEqual("Test3", ws.Cell(2, 10).Value);
            Assert.AreEqual(new DateTime(2017, 11, 3), ws.Cell(2, 11).Value);
            Assert.AreEqual(5500.80, ws.Cell(2, 12).Value);
            Assert.AreEqual("1533_5456", ws.Cell(2, 13).Value);
            Assert.AreEqual(1533d, ws.Cell(3, 10).Value);
            Assert.AreEqual(5456d, ws.Cell(3, 11).Value);
            Assert.AreEqual("String parameter", ws.Cell(3, 12).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 10).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 10).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 10).Style.Border.LeftBorder);
            Assert.AreEqual(XLCellValues.Number, ws.Cell(2, 12).DataType);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 14).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 14).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 14).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 12).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 12).Value);

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderIEnumerableHorizontalNoShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

            ws.Cell(2, 4).DataType = XLCellValues.Number;

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";
            ws.Cell(2, 4).Value = "{di:Sum}";
            ws.Cell(2, 5).Value = "{di:Contacts}";
            ws.Cell(3, 2).Value = "{di:Contacts.Phone}";
            ws.Cell(3, 3).Value = "{di:Contacts.Fax}";
            ws.Cell(3, 4).Value = "{p:StrParam}";

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(4, 1).Value = "{di:Name}";
            ws.Cell(1, 6).Value = "{di:Name}";
            ws.Cell(4, 6).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 6).Value = "{di:Name}";
            ws.Cell(1, 4).Value = "{di:Name}";
            ws.Cell(4, 4).Value = "{di:Name}";
            ws.Cell(2, 14).Value = "{di:Date}";

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("TestRange"), report)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.NoShift,
            };
            panel.Render();

            Assert.AreEqual(29, ws.CellsUsed().Count());
            Assert.AreEqual("Test1", ws.Cell(2, 2).Value);
            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(2, 3).Value);
            Assert.AreEqual(55.76, ws.Cell(2, 4).Value);
            Assert.AreEqual("15_345", ws.Cell(2, 5).Value);
            Assert.AreEqual(15d, ws.Cell(3, 2).Value);
            Assert.AreEqual(345d, ws.Cell(3, 3).Value);
            Assert.AreEqual("String parameter", ws.Cell(3, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLCellValues.Number, ws.Cell(2, 4).DataType);

            Assert.AreEqual("Test2", ws.Cell(2, 6).Value);
            Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(2, 7).Value);
            Assert.AreEqual(110d, ws.Cell(2, 8).Value);
            Assert.AreEqual("76_753465", ws.Cell(2, 9).Value);
            Assert.AreEqual(76d, ws.Cell(3, 6).Value);
            Assert.AreEqual(753465d, ws.Cell(3, 7).Value);
            Assert.AreEqual("String parameter", ws.Cell(3, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.LeftBorder);
            Assert.AreEqual(XLCellValues.Number, ws.Cell(2, 8).DataType);

            Assert.AreEqual("Test3", ws.Cell(2, 10).Value);
            Assert.AreEqual(new DateTime(2017, 11, 3), ws.Cell(2, 11).Value);
            Assert.AreEqual(5500.80, ws.Cell(2, 12).Value);
            Assert.AreEqual("1533_5456", ws.Cell(2, 13).Value);
            Assert.AreEqual(1533d, ws.Cell(3, 10).Value);
            Assert.AreEqual(5456d, ws.Cell(3, 11).Value);
            Assert.AreEqual("String parameter", ws.Cell(3, 12).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 10).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 10).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 10).Style.Border.LeftBorder);
            Assert.AreEqual(XLCellValues.Number, ws.Cell(2, 12).DataType);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 4).Value);
            Assert.AreEqual("{di:Date}", ws.Cell(2, 14).Value);

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}