using System;
using System.Linq;
using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelReporter.Enums;
using ExcelReporter.Excel;

namespace ExcelReporter.Tests.Excel
{
    [TestClass]
    public class ExcelHelperTest
    {
        [TestMethod]
        public void TestIsCellInsideRange()
        {
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLRange range = ws.Range(3, 3, 7, 7);
            Assert.IsTrue(ExcelHelper.IsCellInsideRange(ws.Cell(3, 3), range));
            Assert.IsTrue(ExcelHelper.IsCellInsideRange(ws.Cell(3, 4), range));
            Assert.IsTrue(ExcelHelper.IsCellInsideRange(ws.Cell(4, 4), range));
            Assert.IsTrue(ExcelHelper.IsCellInsideRange(ws.Cell(7, 7), range));
            Assert.IsFalse(ExcelHelper.IsCellInsideRange(ws.Cell(7, 8), range));
            Assert.IsFalse(ExcelHelper.IsCellInsideRange(ws.Cell(2, 3), range));

            range = ws.Range(3, 3, 3, 3);
            Assert.IsTrue(ExcelHelper.IsCellInsideRange(ws.Cell(3, 3), range));
            Assert.IsFalse(ExcelHelper.IsCellInsideRange(ws.Cell(3, 4), range));
            Assert.IsFalse(ExcelHelper.IsCellInsideRange(ws.Cell(2, 3), range));

            IXLWorksheet ws2 = wb.AddWorksheet("Test2");
            range = ws2.Range(3, 3, 3, 3);
            Assert.IsTrue(ExcelHelper.IsCellInsideRange(ws2.Cell(3, 3), range));
            Assert.IsFalse(ExcelHelper.IsCellInsideRange(ws.Cell(3, 3), range));
            Assert.IsFalse(ExcelHelper.IsCellInsideRange(ws2.Cell(3, 4), range));
            Assert.IsFalse(ExcelHelper.IsCellInsideRange(ws.Cell(3, 4), range));
        }

        [TestMethod]
        public void TestIsRangeInsideAnotherRange()
        {
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLRange parent = ws.Range(3, 3, 7, 7);
            Assert.IsTrue(ExcelHelper.IsRangeInsideAnotherRange(parent, ws.Range(3, 3, 7, 7)));
            Assert.IsTrue(ExcelHelper.IsRangeInsideAnotherRange(parent, ws.Range(3, 3, 7, 6)));
            Assert.IsTrue(ExcelHelper.IsRangeInsideAnotherRange(parent, ws.Range(4, 4, 5, 5)));
            Assert.IsTrue(ExcelHelper.IsRangeInsideAnotherRange(parent, ws.Range(4, 4, 4, 4)));
            Assert.IsTrue(ExcelHelper.IsRangeInsideAnotherRange(parent, ws.Range(7, 7, 7, 7)));
            Assert.IsFalse(ExcelHelper.IsRangeInsideAnotherRange(parent, ws.Range(2, 3, 7, 7)));
            Assert.IsFalse(ExcelHelper.IsRangeInsideAnotherRange(parent, ws.Range(7, 7, 7, 8)));

            IXLWorksheet ws2 = wb.AddWorksheet("Test2");
            parent = ws2.Range(3, 3, 7, 7);
            Assert.IsTrue(ExcelHelper.IsRangeInsideAnotherRange(parent, ws2.Range(3, 3, 7, 7)));
            Assert.IsFalse(ExcelHelper.IsRangeInsideAnotherRange(parent, ws.Range(3, 3, 7, 7)));
            Assert.IsFalse(ExcelHelper.IsRangeInsideAnotherRange(parent, ws2.Range(2, 3, 7, 7)));
            Assert.IsFalse(ExcelHelper.IsRangeInsideAnotherRange(parent, ws.Range(2, 3, 7, 7)));
        }

        [TestMethod]
        public void TestGetNearestParentRange()
        {
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLRange parentRange1 = ws.Range(1, 1, 7, 7);
            IXLRange childRange = ws.Range(7, 7, 7, 7);
            Assert.AreEqual(parentRange1, ExcelHelper.GetNearestParentRange(new[] { parentRange1 }, childRange));

            IXLRange parentRange2 = ws.Range(2, 2, 7, 7);
            Assert.AreEqual(parentRange2, ExcelHelper.GetNearestParentRange(new[] { parentRange1, parentRange2 }, childRange));

            IXLRange notParentRange = ws.Range(2, 2, 7, 6);
            Assert.AreEqual(parentRange2, ExcelHelper.GetNearestParentRange(new[] { parentRange1, parentRange2, notParentRange }, childRange));

            IXLRange parentRange3 = ws.Range(7, 7, 7, 7);
            Assert.AreEqual(parentRange3, ExcelHelper.GetNearestParentRange(new[] { parentRange1, parentRange2, notParentRange, parentRange3 }, childRange));

            IXLRange parentRange4 = ws.Range(7, 7, 7, 7);
            MyAssert.Throws<InvalidOperationException>(() => ExcelHelper.GetNearestParentRange(new[] { parentRange1, parentRange2, notParentRange, parentRange3, parentRange4 }, childRange),
                "Found more than one nearest parent ranges");

            IXLRange badChildRange = ws.Range(7, 7, 7, 8);
            MyAssert.Throws<InvalidOperationException>(() => ExcelHelper.GetNearestParentRange(new[] { parentRange1, parentRange2, notParentRange, parentRange3, parentRange4 }, badChildRange),
                "Nearest parent range was not found");
        }

        [TestMethod]
        public void TestGetCellCoordsRelativeRange()
        {
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLRange range = ws.Range(3, 3, 6, 7);
            Assert.AreEqual(new CellCoords(1, 1), ExcelHelper.GetCellCoordsRelativeRange(range, ws.Cell(3, 3)));
            Assert.AreEqual(new CellCoords(1, 2), ExcelHelper.GetCellCoordsRelativeRange(range, ws.Cell(3, 4)));
            Assert.AreEqual(new CellCoords(1, 5), ExcelHelper.GetCellCoordsRelativeRange(range, ws.Cell(3, 7)));
            Assert.AreEqual(new CellCoords(2, 1), ExcelHelper.GetCellCoordsRelativeRange(range, ws.Cell(4, 3)));
            Assert.AreEqual(new CellCoords(2, 4), ExcelHelper.GetCellCoordsRelativeRange(range, ws.Cell(4, 6)));
            Assert.AreEqual(new CellCoords(4, 3), ExcelHelper.GetCellCoordsRelativeRange(range, ws.Cell(6, 5)));
            Assert.AreEqual(new CellCoords(4, 5), ExcelHelper.GetCellCoordsRelativeRange(range, ws.Cell(6, 7)));

            IXLCell cell = ws.Cell(1, 1);
            MyAssert.Throws<InvalidOperationException>(() => ExcelHelper.GetCellCoordsRelativeRange(range, cell), $"{range} is not a parent of {cell}");
        }

        [TestMethod]
        public void TestGetRangeCoordsRelativeParent()
        {
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLRange parentRange = ws.Range(3, 3, 6, 7);
            IXLRange childRange = ws.Range(5, 6, 5, 7);
            Assert.AreEqual(new RangeCoords(new CellCoords(3, 4), new CellCoords(3, 5)), ExcelHelper.GetRangeCoordsRelativeParent(parentRange, childRange));

            childRange = ws.Range(3, 3, 6, 7);
            Assert.AreEqual(new RangeCoords(new CellCoords(1, 1), new CellCoords(4, 5)), ExcelHelper.GetRangeCoordsRelativeParent(parentRange, childRange));

            childRange = ws.Range(3, 3, 3, 3);
            Assert.AreEqual(new RangeCoords(new CellCoords(1, 1), new CellCoords(1, 1)), ExcelHelper.GetRangeCoordsRelativeParent(parentRange, childRange));

            childRange = ws.Range(6, 7, 6, 7);
            Assert.AreEqual(new RangeCoords(new CellCoords(4, 5), new CellCoords(4, 5)), ExcelHelper.GetRangeCoordsRelativeParent(parentRange, childRange));

            childRange = ws.Range(4, 4, 6, 5);
            Assert.AreEqual(new RangeCoords(new CellCoords(2, 2), new CellCoords(4, 3)), ExcelHelper.GetRangeCoordsRelativeParent(parentRange, childRange));

            childRange = ws.Range(4, 4, 7, 5);
            MyAssert.Throws<InvalidOperationException>(() => ExcelHelper.GetRangeCoordsRelativeParent(parentRange, childRange), $"{parentRange} is not a parent of {childRange}");
        }

        [TestMethod]
        public void TestGetAddressShift()
        {
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLCell cell1 = ws.Cell(1, 1);
            IXLCell cell2 = ws.Cell(5, 4);
            Assert.AreEqual(new AddressShift(4, 3), ExcelHelper.GetAddressShift(cell2.Address, cell1.Address));
            Assert.AreEqual(new AddressShift(-4, -3), ExcelHelper.GetAddressShift(cell1.Address, cell2.Address));
        }

        [TestMethod]
        public void TestShiftCell()
        {
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLCell cell = ws.Cell(1, 1);
            var shift = new AddressShift(4, 3);

            Assert.AreEqual(ws.Cell(5, 4), ExcelHelper.ShiftCell(cell, shift));

            cell = ws.Cell(5, 4);
            shift = new AddressShift(-4, -3);
            Assert.AreEqual(ws.Cell(1, 1), ExcelHelper.ShiftCell(cell, shift));
        }

        [TestMethod]
        public void TestCopyRange()
        {
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");
            IXLRange range = ws.Range(6, 5, 9, 7);

            range.FirstCell().Value = "RangeStart";
            range.LastCell().Value = "RangeEnd";
            range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.LastCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);

            ws.Cell(11, 9).Value = DateTime.Now;
            ws.Cell(11, 9).DataType = XLCellValues.DateTime;

            IXLRange newRange = ExcelHelper.CopyRange(range, ws.Cell(10, 8));

            Assert.AreEqual(4, ws.CellsUsed().Count());
            Assert.AreEqual(range.FirstCell(), ws.Cell(6, 5));
            Assert.AreEqual(range.LastCell(), ws.Cell(9, 7));
            Assert.AreEqual("RangeStart", range.FirstCell().Value.ToString());
            Assert.AreEqual("RangeEnd", range.LastCell().Value.ToString());
            Assert.AreEqual(XLBorderStyleValues.Thin, range.FirstCell().Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, range.LastCell().Style.Border.BottomBorder);
            Assert.AreEqual(newRange.FirstCell(), ws.Cell(10, 8));
            Assert.AreEqual(newRange.LastCell(), ws.Cell(13, 10));
            Assert.AreEqual("RangeStart", newRange.FirstCell().Value.ToString());
            Assert.AreEqual("RangeEnd", newRange.LastCell().Value.ToString());
            Assert.AreEqual(XLBorderStyleValues.Thin, newRange.FirstCell().Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, newRange.LastCell().Style.Border.BottomBorder);
            Assert.AreEqual(XLCellValues.Text, ws.Cell(11, 9).DataType);

            range.Clear();
            IXLRange newRange2 = ExcelHelper.CopyRange(newRange, ws.Cell(11, 8));

            Assert.AreEqual(3, ws.CellsUsed().Count());
            Assert.AreEqual(newRange.FirstCell(), ws.Cell(10, 8));
            Assert.AreEqual(newRange.LastCell(), ws.Cell(13, 10));
            Assert.AreEqual("RangeStart", newRange.FirstCell().Value.ToString());
            Assert.AreEqual(string.Empty, newRange.LastCell().Value.ToString());
            Assert.AreEqual(XLBorderStyleValues.Thin, newRange.FirstCell().Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, newRange.LastCell().Style.Border.BottomBorder);
            Assert.AreEqual(newRange2.FirstCell(), ws.Cell(11, 8));
            Assert.AreEqual(newRange2.LastCell(), ws.Cell(14, 10));
            Assert.AreEqual("RangeStart", newRange2.FirstCell().Value.ToString());
            Assert.AreEqual("RangeEnd", newRange2.LastCell().Value.ToString());
            Assert.AreEqual(XLBorderStyleValues.Thin, newRange2.FirstCell().Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, newRange2.LastCell().Style.Border.BottomBorder);

            //wb.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestCopyNamedRange()
        {
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");
            IXLRange range = ws.Range(6, 5, 9, 7);
            range.AddToNamed("TestRange", XLScope.Worksheet);
            IXLNamedRange namedRange = ws.NamedRange("TestRange");

            range.FirstCell().Value = "RangeStart";
            range.LastCell().Value = "RangeEnd";
            range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.LastCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);

            ws.Cell(11, 9).Value = DateTime.Now;
            ws.Cell(11, 9).DataType = XLCellValues.DateTime;

            IXLNamedRange copiedNamedRange = ExcelHelper.CopyNamedRange(namedRange, ws.Cell(10, 8), "Copy");
            IXLRange newRange = copiedNamedRange.Ranges.ElementAt(0);

            Assert.AreEqual(1, copiedNamedRange.Ranges.Count);
            Assert.AreEqual(4, ws.CellsUsed().Count());
            Assert.AreEqual(range.FirstCell(), ws.Cell(6, 5));
            Assert.AreEqual(range.LastCell(), ws.Cell(9, 7));
            Assert.AreEqual("RangeStart", range.FirstCell().Value.ToString());
            Assert.AreEqual("RangeEnd", range.LastCell().Value.ToString());
            Assert.AreEqual(XLBorderStyleValues.Thin, range.FirstCell().Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, range.LastCell().Style.Border.BottomBorder);
            Assert.AreEqual(newRange.FirstCell(), ws.Cell(10, 8));
            Assert.AreEqual(newRange.LastCell(), ws.Cell(13, 10));
            Assert.AreEqual("RangeStart", newRange.FirstCell().Value.ToString());
            Assert.AreEqual("RangeEnd", newRange.LastCell().Value.ToString());
            Assert.AreEqual(XLBorderStyleValues.Thin, newRange.FirstCell().Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, newRange.LastCell().Style.Border.BottomBorder);
            Assert.AreEqual(XLCellValues.Text, ws.Cell(11, 9).DataType);

            range.Clear();
            IXLNamedRange copiedNamedRange2 = ExcelHelper.CopyNamedRange(copiedNamedRange, ws.Cell(11, 8), "Copy2");
            IXLRange newRange2 = copiedNamedRange2.Ranges.ElementAt(0);

            Assert.AreEqual(1, copiedNamedRange2.Ranges.Count);
            Assert.AreEqual(3, ws.CellsUsed().Count());
            Assert.AreEqual(newRange.FirstCell(), ws.Cell(10, 8));
            Assert.AreEqual(newRange.LastCell(), ws.Cell(13, 10));
            Assert.AreEqual("RangeStart", newRange.FirstCell().Value.ToString());
            Assert.AreEqual(string.Empty, newRange.LastCell().Value.ToString());
            Assert.AreEqual(XLBorderStyleValues.Thin, newRange.FirstCell().Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, newRange.LastCell().Style.Border.BottomBorder);
            Assert.AreEqual(newRange2.FirstCell(), ws.Cell(11, 8));
            Assert.AreEqual(newRange2.LastCell(), ws.Cell(14, 10));
            Assert.AreEqual("RangeStart", newRange2.FirstCell().Value.ToString());
            Assert.AreEqual("RangeEnd", newRange2.LastCell().Value.ToString());
            Assert.AreEqual(XLBorderStyleValues.Thin, newRange2.FirstCell().Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, newRange2.LastCell().Style.Border.BottomBorder);

            Assert.AreEqual(3, ws.NamedRanges.Count());

            //wb.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestAllocateSpaceForNextRange()
        {
            // Добавление ячеек сверху
            XLWorkbook wb = InitWorkBookForShiftTests();
            IXLWorksheet ws = wb.Worksheet("Test");
            IXLRange range = ws.NamedRange("TestRange").Ranges.ElementAt(0);

            ExcelHelper.AllocateSpaceForNextRange(range, Direction.Top);

            IXLCell rangeStartCell = ws.Cells().Single(c => c.Value.ToString() == "RangeStart");
            IXLCell rangeEndCell = ws.Cells().Single(c => c.Value.ToString() == "RangeEnd");
            IXLCell belowCell1 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_1");
            IXLCell belowCell2 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_2");
            IXLCell rightCell1 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_1");
            IXLCell rightCell2 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_2");
            IXLCell aboveCell1 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_1");
            IXLCell aboveCell2 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_2");
            IXLCell leftCell1 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_1");
            IXLCell leftCell2 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_2");

            Assert.AreEqual(10, ws.CellsUsed().Count());
            Assert.AreEqual(rangeStartCell, ws.Cell(10, 5));
            Assert.AreEqual(rangeEndCell, ws.Cell(13, 7));
            Assert.AreEqual(belowCell1, ws.Cell(14, 6));
            Assert.AreEqual(belowCell2, ws.Cell(10, 8));
            Assert.AreEqual(rightCell1, ws.Cell(7, 8));
            Assert.AreEqual(rightCell2, ws.Cell(5, 8));
            Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
            Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
            Assert.AreEqual(leftCell1, ws.Cell(7, 4));
            Assert.AreEqual(leftCell2, ws.Cell(10, 4));


            // Добавление ячеек снизу
            wb = InitWorkBookForShiftTests();
            ws = wb.Worksheet("Test");
            range = ws.NamedRange("TestRange").Ranges.ElementAt(0);

            ExcelHelper.AllocateSpaceForNextRange(range, Direction.Bottom);

            rangeStartCell = ws.Cells().Single(c => c.Value.ToString() == "RangeStart");
            rangeEndCell = ws.Cells().Single(c => c.Value.ToString() == "RangeEnd");
            belowCell1 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_1");
            belowCell2 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_2");
            rightCell1 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_1");
            rightCell2 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_2");
            aboveCell1 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_1");
            aboveCell2 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_2");
            leftCell1 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_1");
            leftCell2 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_2");

            Assert.AreEqual(10, ws.CellsUsed().Count());
            Assert.AreEqual(rangeStartCell, ws.Cell(6, 5));
            Assert.AreEqual(rangeEndCell, ws.Cell(9, 7));
            Assert.AreEqual(belowCell1, ws.Cell(14, 6));
            Assert.AreEqual(belowCell2, ws.Cell(10, 8));
            Assert.AreEqual(rightCell1, ws.Cell(7, 8));
            Assert.AreEqual(rightCell2, ws.Cell(5, 8));
            Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
            Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
            Assert.AreEqual(leftCell1, ws.Cell(7, 4));
            Assert.AreEqual(leftCell2, ws.Cell(10, 4));


            // Добавление строк сверху
            wb = InitWorkBookForShiftTests();
            ws = wb.Worksheet("Test");
            range = ws.NamedRange("TestRange").Ranges.ElementAt(0);

            ExcelHelper.AllocateSpaceForNextRange(range, Direction.Top, ShiftType.Row);

            rangeStartCell = ws.Cells().Single(c => c.Value.ToString() == "RangeStart");
            rangeEndCell = ws.Cells().Single(c => c.Value.ToString() == "RangeEnd");
            belowCell1 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_1");
            belowCell2 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_2");
            rightCell1 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_1");
            rightCell2 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_2");
            aboveCell1 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_1");
            aboveCell2 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_2");
            leftCell1 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_1");
            leftCell2 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_2");

            Assert.AreEqual(10, ws.CellsUsed().Count());
            Assert.AreEqual(rangeStartCell, ws.Cell(10, 5));
            Assert.AreEqual(rangeEndCell, ws.Cell(13, 7));
            Assert.AreEqual(belowCell1, ws.Cell(14, 6));
            Assert.AreEqual(belowCell2, ws.Cell(14, 8));
            Assert.AreEqual(rightCell1, ws.Cell(11, 8));
            Assert.AreEqual(rightCell2, ws.Cell(5, 8));
            Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
            Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
            Assert.AreEqual(leftCell1, ws.Cell(11, 4));
            Assert.AreEqual(leftCell2, ws.Cell(14, 4));


            // Добавление строк снизу
            wb = InitWorkBookForShiftTests();
            ws = wb.Worksheet("Test");
            range = ws.NamedRange("TestRange").Ranges.ElementAt(0);

            ExcelHelper.AllocateSpaceForNextRange(range, Direction.Bottom, ShiftType.Row);

            rangeStartCell = ws.Cells().Single(c => c.Value.ToString() == "RangeStart");
            rangeEndCell = ws.Cells().Single(c => c.Value.ToString() == "RangeEnd");
            belowCell1 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_1");
            belowCell2 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_2");
            rightCell1 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_1");
            rightCell2 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_2");
            aboveCell1 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_1");
            aboveCell2 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_2");
            leftCell1 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_1");
            leftCell2 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_2");

            Assert.AreEqual(10, ws.CellsUsed().Count());
            Assert.AreEqual(rangeStartCell, ws.Cell(6, 5));
            Assert.AreEqual(rangeEndCell, ws.Cell(9, 7));
            Assert.AreEqual(belowCell1, ws.Cell(14, 6));
            Assert.AreEqual(belowCell2, ws.Cell(14, 8));
            Assert.AreEqual(rightCell1, ws.Cell(7, 8));
            Assert.AreEqual(rightCell2, ws.Cell(5, 8));
            Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
            Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
            Assert.AreEqual(leftCell1, ws.Cell(7, 4));
            Assert.AreEqual(leftCell2, ws.Cell(14, 4));


            // Добавление ячеек слева
            wb = InitWorkBookForShiftTests();
            ws = wb.Worksheet("Test");
            range = ws.NamedRange("TestRange").Ranges.ElementAt(0);

            ExcelHelper.AllocateSpaceForNextRange(range, Direction.Left);

            rangeStartCell = ws.Cells().Single(c => c.Value.ToString() == "RangeStart");
            rangeEndCell = ws.Cells().Single(c => c.Value.ToString() == "RangeEnd");
            belowCell1 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_1");
            belowCell2 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_2");
            rightCell1 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_1");
            rightCell2 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_2");
            aboveCell1 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_1");
            aboveCell2 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_2");
            leftCell1 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_1");
            leftCell2 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_2");

            Assert.AreEqual(10, ws.CellsUsed().Count());
            Assert.AreEqual(rangeStartCell, ws.Cell(6, 8));
            Assert.AreEqual(rangeEndCell, ws.Cell(9, 10));
            Assert.AreEqual(belowCell1, ws.Cell(10, 6));
            Assert.AreEqual(belowCell2, ws.Cell(10, 8));
            Assert.AreEqual(rightCell1, ws.Cell(7, 11));
            Assert.AreEqual(rightCell2, ws.Cell(5, 8));
            Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
            Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
            Assert.AreEqual(leftCell1, ws.Cell(7, 4));
            Assert.AreEqual(leftCell2, ws.Cell(10, 4));


            // Добавление ячеек справа
            wb = InitWorkBookForShiftTests();
            ws = wb.Worksheet("Test");
            range = ws.NamedRange("TestRange").Ranges.ElementAt(0);

            ExcelHelper.AllocateSpaceForNextRange(range, Direction.Right);

            rangeStartCell = ws.Cells().Single(c => c.Value.ToString() == "RangeStart");
            rangeEndCell = ws.Cells().Single(c => c.Value.ToString() == "RangeEnd");
            belowCell1 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_1");
            belowCell2 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_2");
            rightCell1 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_1");
            rightCell2 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_2");
            aboveCell1 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_1");
            aboveCell2 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_2");
            leftCell1 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_1");
            leftCell2 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_2");

            Assert.AreEqual(10, ws.CellsUsed().Count());
            Assert.AreEqual(rangeStartCell, ws.Cell(6, 5));
            Assert.AreEqual(rangeEndCell, ws.Cell(9, 7));
            Assert.AreEqual(belowCell1, ws.Cell(10, 6));
            Assert.AreEqual(belowCell2, ws.Cell(10, 8));
            Assert.AreEqual(rightCell1, ws.Cell(7, 11));
            Assert.AreEqual(rightCell2, ws.Cell(5, 8));
            Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
            Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
            Assert.AreEqual(leftCell1, ws.Cell(7, 4));
            Assert.AreEqual(leftCell2, ws.Cell(10, 4));


            // Добавление колонок слева
            wb = InitWorkBookForShiftTests();
            ws = wb.Worksheet("Test");
            range = ws.NamedRange("TestRange").Ranges.ElementAt(0);

            ExcelHelper.AllocateSpaceForNextRange(range, Direction.Left, ShiftType.Row);

            rangeStartCell = ws.Cells().Single(c => c.Value.ToString() == "RangeStart");
            rangeEndCell = ws.Cells().Single(c => c.Value.ToString() == "RangeEnd");
            belowCell1 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_1");
            belowCell2 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_2");
            rightCell1 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_1");
            rightCell2 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_2");
            aboveCell1 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_1");
            aboveCell2 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_2");
            leftCell1 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_1");
            leftCell2 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_2");

            Assert.AreEqual(10, ws.CellsUsed().Count());
            Assert.AreEqual(rangeStartCell, ws.Cell(6, 8));
            Assert.AreEqual(rangeEndCell, ws.Cell(9, 10));
            Assert.AreEqual(belowCell1, ws.Cell(10, 9));
            Assert.AreEqual(belowCell2, ws.Cell(10, 11));
            Assert.AreEqual(rightCell1, ws.Cell(7, 11));
            Assert.AreEqual(rightCell2, ws.Cell(5, 11));
            Assert.AreEqual(aboveCell1, ws.Cell(5, 9));
            Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
            Assert.AreEqual(leftCell1, ws.Cell(7, 4));
            Assert.AreEqual(leftCell2, ws.Cell(10, 4));


            // Добавление колонок справа
            wb = InitWorkBookForShiftTests();
            ws = wb.Worksheet("Test");
            range = ws.NamedRange("TestRange").Ranges.ElementAt(0);

            ExcelHelper.AllocateSpaceForNextRange(range, Direction.Right, ShiftType.Row);

            rangeStartCell = ws.Cells().Single(c => c.Value.ToString() == "RangeStart");
            rangeEndCell = ws.Cells().Single(c => c.Value.ToString() == "RangeEnd");
            belowCell1 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_1");
            belowCell2 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_2");
            rightCell1 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_1");
            rightCell2 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_2");
            aboveCell1 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_1");
            aboveCell2 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_2");
            leftCell1 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_1");
            leftCell2 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_2");

            Assert.AreEqual(10, ws.CellsUsed().Count());
            Assert.AreEqual(rangeStartCell, ws.Cell(6, 5));
            Assert.AreEqual(rangeEndCell, ws.Cell(9, 7));
            Assert.AreEqual(belowCell1, ws.Cell(10, 6));
            Assert.AreEqual(belowCell2, ws.Cell(10, 11));
            Assert.AreEqual(rightCell1, ws.Cell(7, 11));
            Assert.AreEqual(rightCell2, ws.Cell(5, 11));
            Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
            Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
            Assert.AreEqual(leftCell1, ws.Cell(7, 4));
            Assert.AreEqual(leftCell2, ws.Cell(10, 4));


            // Ничего не добавляется (без сдвига)
            wb = InitWorkBookForShiftTests();
            ws = wb.Worksheet("Test");
            range = ws.NamedRange("TestRange").Ranges.ElementAt(0);

            ExcelHelper.AllocateSpaceForNextRange(range, Direction.Top, ShiftType.NoShift);

            rangeStartCell = ws.Cells().Single(c => c.Value.ToString() == "RangeStart");
            rangeEndCell = ws.Cells().Single(c => c.Value.ToString() == "RangeEnd");
            belowCell1 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_1");
            belowCell2 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_2");
            rightCell1 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_1");
            rightCell2 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_2");
            aboveCell1 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_1");
            aboveCell2 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_2");
            leftCell1 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_1");
            leftCell2 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_2");

            Assert.AreEqual(10, ws.CellsUsed().Count());
            Assert.AreEqual(rangeStartCell, ws.Cell(6, 5));
            Assert.AreEqual(rangeEndCell, ws.Cell(9, 7));
            Assert.AreEqual(belowCell1, ws.Cell(10, 6));
            Assert.AreEqual(belowCell2, ws.Cell(10, 8));
            Assert.AreEqual(rightCell1, ws.Cell(7, 8));
            Assert.AreEqual(rightCell2, ws.Cell(5, 8));
            Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
            Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
            Assert.AreEqual(leftCell1, ws.Cell(7, 4));
            Assert.AreEqual(leftCell2, ws.Cell(10, 4));

            //wb.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestDeleteRange()
        {
            // Удаление со сдвигом ячеек вверх
            XLWorkbook wb = InitWorkBookForShiftTests();
            IXLWorksheet ws = wb.Worksheet("Test");
            IXLRange range = ws.NamedRange("TestRange").Ranges.ElementAt(0);

            ExcelHelper.DeleteRange(range, ShiftType.Cells);

            IXLCell rangeStartCell = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RangeStart");
            IXLCell rangeEndCell = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RangeEnd");
            IXLCell belowCell1 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_1");
            IXLCell belowCell2 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_2");
            IXLCell rightCell1 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_1");
            IXLCell rightCell2 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_2");
            IXLCell aboveCell1 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_1");
            IXLCell aboveCell2 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_2");
            IXLCell leftCell1 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_1");
            IXLCell leftCell2 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_2");

            Assert.IsNull(rangeStartCell);
            Assert.IsNull(rangeEndCell);
            Assert.AreEqual(8, ws.CellsUsed().Count());
            Assert.AreEqual(belowCell1, ws.Cell(6, 6));
            Assert.AreEqual(belowCell2, ws.Cell(10, 8));
            Assert.AreEqual(rightCell1, ws.Cell(7, 8));
            Assert.AreEqual(rightCell2, ws.Cell(5, 8));
            Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
            Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
            Assert.AreEqual(leftCell1, ws.Cell(7, 4));
            Assert.AreEqual(leftCell2, ws.Cell(10, 4));

            // Удаление со сдвигом строки вверх
            wb = InitWorkBookForShiftTests();
            ws = wb.Worksheet("Test");
            range = ws.NamedRange("TestRange").Ranges.ElementAt(0);

            ExcelHelper.DeleteRange(range, ShiftType.Row);

            rangeStartCell = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RangeStart");
            rangeEndCell = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RangeEnd");
            belowCell1 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_1");
            belowCell2 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_2");
            rightCell1 = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RightCell_1");
            rightCell2 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_2");
            aboveCell1 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_1");
            aboveCell2 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_2");
            leftCell1 = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "LeftCell_1");
            leftCell2 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_2");

            Assert.IsNull(rangeStartCell);
            Assert.IsNull(rangeEndCell);
            Assert.IsNull(leftCell1);
            Assert.IsNull(rightCell1);
            Assert.AreEqual(6, ws.CellsUsed().Count());
            Assert.AreEqual(belowCell1, ws.Cell(6, 6));
            Assert.AreEqual(belowCell2, ws.Cell(6, 8));
            Assert.AreEqual(rightCell2, ws.Cell(5, 8));
            Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
            Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
            Assert.AreEqual(leftCell2, ws.Cell(6, 4));

            // Удаление со сдвигом ячеек влево
            wb = InitWorkBookForShiftTests();
            ws = wb.Worksheet("Test");
            range = ws.NamedRange("TestRange").Ranges.ElementAt(0);

            ExcelHelper.DeleteRange(range, ShiftType.Cells, XLShiftDeletedCells.ShiftCellsLeft);

            rangeStartCell = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RangeStart");
            rangeEndCell = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RangeEnd");
            belowCell1 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_1");
            belowCell2 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_2");
            rightCell1 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_1");
            rightCell2 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_2");
            aboveCell1 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_1");
            aboveCell2 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_2");
            leftCell1 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_1");
            leftCell2 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_2");

            Assert.IsNull(rangeStartCell);
            Assert.IsNull(rangeEndCell);
            Assert.AreEqual(8, ws.CellsUsed().Count());
            Assert.AreEqual(belowCell1, ws.Cell(10, 6));
            Assert.AreEqual(belowCell2, ws.Cell(10, 8));
            Assert.AreEqual(rightCell1, ws.Cell(7, 5));
            Assert.AreEqual(rightCell2, ws.Cell(5, 8));
            Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
            Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
            Assert.AreEqual(leftCell1, ws.Cell(7, 4));
            Assert.AreEqual(leftCell2, ws.Cell(10, 4));

            // Удаление со сдвигом колонки влево
            wb = InitWorkBookForShiftTests();
            ws = wb.Worksheet("Test");
            range = ws.NamedRange("TestRange").Ranges.ElementAt(0);

            ExcelHelper.DeleteRange(range, ShiftType.Row, XLShiftDeletedCells.ShiftCellsLeft);

            rangeStartCell = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RangeStart");
            rangeEndCell = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RangeEnd");
            belowCell1 = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "BelowCell_1");
            belowCell2 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_2");
            rightCell1 = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RightCell_1");
            rightCell2 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_2");
            aboveCell1 = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "AboveCell_1");
            aboveCell2 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_2");
            leftCell1 = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "LeftCell_1");
            leftCell2 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_2");

            Assert.IsNull(rangeStartCell);
            Assert.IsNull(rangeEndCell);
            Assert.IsNull(aboveCell1);
            Assert.IsNull(belowCell1);
            Assert.AreEqual(6, ws.CellsUsed().Count());
            Assert.AreEqual(belowCell2, ws.Cell(10, 5));
            Assert.AreEqual(rightCell1, ws.Cell(7, 5));
            Assert.AreEqual(rightCell2, ws.Cell(5, 5));
            Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
            Assert.AreEqual(leftCell1, ws.Cell(7, 4));
            Assert.AreEqual(leftCell2, ws.Cell(10, 4));

            // Удаление без сдвига
            wb = InitWorkBookForShiftTests();
            ws = wb.Worksheet("Test");
            range = ws.NamedRange("TestRange").Ranges.ElementAt(0);

            ExcelHelper.DeleteRange(range, ShiftType.NoShift);

            rangeStartCell = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RangeStart");
            rangeEndCell = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RangeEnd");
            belowCell1 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_1");
            belowCell2 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_2");
            rightCell1 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_1");
            rightCell2 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_2");
            aboveCell1 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_1");
            aboveCell2 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_2");
            leftCell1 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_1");
            leftCell2 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_2");

            Assert.IsNull(rangeStartCell);
            Assert.IsNull(rangeEndCell);
            Assert.AreEqual(XLBorderStyleValues.None, range.FirstCell().Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, range.LastCell().Style.Border.BottomBorder);
            Assert.AreEqual(8, ws.CellsUsed().Count());
            Assert.AreEqual(belowCell1, ws.Cell(10, 6));
            Assert.AreEqual(belowCell2, ws.Cell(10, 8));
            Assert.AreEqual(rightCell1, ws.Cell(7, 8));
            Assert.AreEqual(rightCell2, ws.Cell(5, 8));
            Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
            Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
            Assert.AreEqual(leftCell1, ws.Cell(7, 4));
            Assert.AreEqual(leftCell2, ws.Cell(10, 4));

            //wb.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestMoveRange()
        {
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");
            IXLRange range = ws.Range(6, 5, 9, 7);

            range.FirstCell().Value = "RangeStart";
            range.LastCell().Value = "RangeEnd";
            range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.LastCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);

            ws.Cell(11, 9).Value = DateTime.Now;
            ws.Cell(11, 9).DataType = XLCellValues.DateTime;

            IXLRange movedRange = ExcelHelper.MoveRange(range, ws.Cell(10, 8));

            Assert.AreEqual(2, ws.CellsUsed().Count());
            Assert.AreEqual(range.FirstCell(), ws.Cell(6, 5));
            Assert.AreEqual(range.LastCell(), ws.Cell(9, 7));
            Assert.AreEqual(string.Empty, range.FirstCell().Value.ToString());
            Assert.AreEqual(string.Empty, range.LastCell().Value.ToString());
            Assert.AreEqual(XLBorderStyleValues.None, range.FirstCell().Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, range.LastCell().Style.Border.BottomBorder);
            Assert.AreEqual(movedRange.FirstCell(), ws.Cell(10, 8));
            Assert.AreEqual(movedRange.LastCell(), ws.Cell(13, 10));
            Assert.AreEqual("RangeStart", movedRange.FirstCell().Value.ToString());
            Assert.AreEqual("RangeEnd", movedRange.LastCell().Value.ToString());
            Assert.AreEqual(XLBorderStyleValues.Thin, movedRange.FirstCell().Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, movedRange.LastCell().Style.Border.BottomBorder);
            Assert.AreEqual(XLCellValues.Text, ws.Cell(11, 9).DataType);

            IXLRange movedRange2 = ExcelHelper.MoveRange(movedRange, ws.Cell(11, 8));

            Assert.AreEqual(2, ws.CellsUsed().Count());
            Assert.AreEqual(movedRange.FirstCell(), ws.Cell(10, 8));
            Assert.AreEqual(movedRange.LastCell(), ws.Cell(13, 10));
            Assert.AreEqual(string.Empty, movedRange.FirstCell().Value.ToString());
            Assert.AreEqual(string.Empty, movedRange.LastCell().Value.ToString());
            Assert.AreEqual(XLBorderStyleValues.None, movedRange.FirstCell().Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, movedRange.LastCell().Style.Border.BottomBorder);
            Assert.AreEqual(movedRange2.FirstCell(), ws.Cell(11, 8));
            Assert.AreEqual(movedRange2.LastCell(), ws.Cell(14, 10));
            Assert.AreEqual("RangeStart", movedRange2.FirstCell().Value.ToString());
            Assert.AreEqual("RangeEnd", movedRange2.LastCell().Value.ToString());
            Assert.AreEqual(XLBorderStyleValues.Thin, movedRange2.FirstCell().Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, movedRange2.LastCell().Style.Border.BottomBorder);

            //wb.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestMoveNamedRange()
        {
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");
            IXLRange range = ws.Range(6, 5, 9, 7);
            range.AddToNamed("TestRange", XLScope.Worksheet);
            IXLNamedRange namedRange = ws.NamedRange("TestRange");

            range.FirstCell().Value = "RangeStart";
            range.LastCell().Value = "RangeEnd";
            range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.LastCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);

            ws.Cell(11, 9).Value = DateTime.Now;
            ws.Cell(11, 9).DataType = XLCellValues.DateTime;

            IXLNamedRange movedNamedRange = ExcelHelper.MoveNamedRange(namedRange, ws.Cell(10, 8));
            IXLRange movedRange = movedNamedRange.Ranges.ElementAt(0);

            Assert.AreEqual(movedNamedRange, ws.NamedRange("TestRange"));
            Assert.AreEqual(1, movedNamedRange.Ranges.Count);
            Assert.AreEqual(2, ws.CellsUsed().Count());
            Assert.AreEqual(range.FirstCell(), ws.Cell(6, 5));
            Assert.AreEqual(range.LastCell(), ws.Cell(9, 7));
            Assert.AreEqual(string.Empty, range.FirstCell().Value.ToString());
            Assert.AreEqual(string.Empty, range.LastCell().Value.ToString());
            Assert.AreEqual(XLBorderStyleValues.None, range.FirstCell().Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, range.LastCell().Style.Border.BottomBorder);
            Assert.AreEqual(movedRange.FirstCell(), ws.Cell(10, 8));
            Assert.AreEqual(movedRange.LastCell(), ws.Cell(13, 10));
            Assert.AreEqual("RangeStart", movedRange.FirstCell().Value.ToString());
            Assert.AreEqual("RangeEnd", movedRange.LastCell().Value.ToString());
            Assert.AreEqual(XLBorderStyleValues.Thin, movedRange.FirstCell().Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, movedRange.LastCell().Style.Border.BottomBorder);
            Assert.AreEqual(XLCellValues.Text, ws.Cell(11, 9).DataType);

            IXLNamedRange movedNamedRange2 = ExcelHelper.MoveNamedRange(movedNamedRange, ws.Cell(11, 8));
            IXLRange movedRange2 = movedNamedRange2.Ranges.ElementAt(0);

            Assert.AreEqual(movedNamedRange2, ws.NamedRange("TestRange"));
            Assert.AreEqual(1, movedNamedRange2.Ranges.Count);
            Assert.AreEqual(2, ws.CellsUsed().Count());
            Assert.AreEqual(movedRange.FirstCell(), ws.Cell(10, 8));
            Assert.AreEqual(movedRange.LastCell(), ws.Cell(13, 10));
            Assert.AreEqual(string.Empty, movedRange.FirstCell().Value.ToString());
            Assert.AreEqual(string.Empty, movedRange.LastCell().Value.ToString());
            Assert.AreEqual(XLBorderStyleValues.None, movedRange.FirstCell().Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, movedRange.LastCell().Style.Border.BottomBorder);
            Assert.AreEqual(movedRange2.FirstCell(), ws.Cell(11, 8));
            Assert.AreEqual(movedRange2.LastCell(), ws.Cell(14, 10));
            Assert.AreEqual("RangeStart", movedRange2.FirstCell().Value.ToString());
            Assert.AreEqual("RangeEnd", movedRange2.LastCell().Value.ToString());
            Assert.AreEqual(XLBorderStyleValues.Thin, movedRange2.FirstCell().Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, movedRange2.LastCell().Style.Border.BottomBorder);

            //wb.SaveAs("test.xlsx");
        }

        private XLWorkbook InitWorkBookForShiftTests()
        {
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLRange range = ws.Range(6, 5, 9, 7);
            range.AddToNamed("TestRange", XLScope.Worksheet);
            range.FirstCell().Value = "RangeStart";
            range.LastCell().Value = "RangeEnd";
            range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.LastCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);

            ws.Cell(10, 6).Value = "BelowCell_1";
            ws.Cell(10, 8).Value = "BelowCell_2";
            ws.Cell(7, 8).Value = "RightCell_1";
            ws.Cell(5, 8).Value = "RightCell_2";
            ws.Cell(5, 6).Value = "AboveCell_1";
            ws.Cell(5, 4).Value = "AboveCell_2";
            ws.Cell(7, 4).Value = "LeftCell_1";
            ws.Cell(10, 4).Value = "LeftCell_2";

            return wb;
        }
    }
}