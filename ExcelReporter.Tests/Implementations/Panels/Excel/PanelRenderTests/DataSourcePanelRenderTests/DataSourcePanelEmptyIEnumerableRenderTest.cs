using System;
using System.Linq;
using ClosedXML.Excel;
using ExcelReporter.Enums;
using ExcelReporter.Implementations.Panels.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Implementations.Panels.Excel.PanelRenderTests.DataSourcePanelRenderTests
{
    [TestClass]
    public class DataSourcePanelEmptyIEnumerableRenderTest
    {
        [TestMethod]
        public void TestRenderEmptyIEnumerableVerticalCellsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            range.Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.Style.Border.SetRightBorder(XLBorderStyleValues.Thin);
            range.Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
            range.Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

            ws.Cell(4, 3).Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
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

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetEmptyIEnumerable()", ws.NamedRange("TestRange"), report);
            panel.Render();

            Assert.AreEqual(8, ws.CellsUsed().Count());
            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(2, 4).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(2, 4).Value);

            Assert.AreEqual(0, ws.Cells().Count(c => c.DataType == XLCellValues.Number));
            Assert.AreEqual(1, ws.Cells().Count(c => c.Style.Border.TopBorder == XLBorderStyleValues.Thin));
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.TopBorder);

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderEmptyIEnumerableVerticalRowsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            range.Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.Style.Border.SetRightBorder(XLBorderStyleValues.Thin);
            range.Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
            range.Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

            ws.Cell(4, 3).Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
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

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetEmptyIEnumerable()", ws.NamedRange("TestRange"), report)
            {
                ShiftType = ShiftType.Row,
            };
            panel.Render();

            Assert.AreEqual(6, ws.CellsUsed().Count());
            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(2, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(2, 4).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(2, 6).Value);

            Assert.AreEqual(0, ws.Cells().Count(c => c.DataType == XLCellValues.Number));
            Assert.AreEqual(1, ws.Cells().Count(c => c.Style.Border.TopBorder == XLBorderStyleValues.Thin));
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.TopBorder);

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderEmptyIEnumerableVerticalNoShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            range.Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.Style.Border.SetRightBorder(XLBorderStyleValues.Thin);
            range.Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
            range.Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

            ws.Cell(4, 3).Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
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

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetEmptyIEnumerable()", ws.NamedRange("TestRange"), report)
            {
                ShiftType = ShiftType.NoShift,
            };
            panel.Render();

            Assert.AreEqual(8, ws.CellsUsed().Count());
            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 4).Value);

            Assert.AreEqual(0, ws.Cells().Count(c => c.DataType == XLCellValues.Number));
            Assert.AreEqual(1, ws.Cells().Count(c => c.Style.Border.TopBorder == XLBorderStyleValues.Thin));
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.TopBorder);

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderEmptyIEnumerableHorizontalCellsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            range.Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.Style.Border.SetRightBorder(XLBorderStyleValues.Thin);
            range.Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
            range.Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

            ws.Cell(2, 6).Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);
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

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetEmptyIEnumerable()", ws.NamedRange("TestRange"), report)
            {
                Type = PanelType.Horizontal,
            };
            panel.Render();

            Assert.AreEqual(8, ws.CellsUsed().Count());
            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 2).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 4).Value);

            Assert.AreEqual(0, ws.Cells().Count(c => c.DataType == XLCellValues.Number));
            Assert.AreEqual(1, ws.Cells().Count(c => c.Style.Border.LeftBorder == XLBorderStyleValues.Thin));
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.LeftBorder);

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderEmptyIEnumerableHorizontalRowsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            range.Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.Style.Border.SetRightBorder(XLBorderStyleValues.Thin);
            range.Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
            range.Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

            ws.Cell(2, 6).Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);
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

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetEmptyIEnumerable()", ws.NamedRange("TestRange"), report)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.Row,
            };
            panel.Render();

            Assert.AreEqual(6, ws.CellsUsed().Count());
            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 2).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 2).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 2).Value);

            Assert.AreEqual(0, ws.Cells().Count(c => c.DataType == XLCellValues.Number));
            Assert.AreEqual(1, ws.Cells().Count(c => c.Style.Border.LeftBorder == XLBorderStyleValues.Thin));
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.LeftBorder);

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderEmptyIEnumerableHorizontalNoShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            range.Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.Style.Border.SetRightBorder(XLBorderStyleValues.Thin);
            range.Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
            range.Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

            ws.Cell(2, 6).Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);
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

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetEmptyIEnumerable()", ws.NamedRange("TestRange"), report)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.NoShift,
            };
            panel.Render();

            Assert.AreEqual(8, ws.CellsUsed().Count());
            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 4).Value);

            Assert.AreEqual(0, ws.Cells().Count(c => c.DataType == XLCellValues.Number));
            Assert.AreEqual(1, ws.Cells().Count(c => c.Style.Border.LeftBorder == XLBorderStyleValues.Thin));
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.LeftBorder);

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}