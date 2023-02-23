using System;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using NUnit.Framework;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Excel;
using ExcelReportGenerator.Tests.CustomAsserts;

namespace ExcelReportGenerator.Tests.Excel
{
    public class ExcelHelperTest
    {
        [Test]
        public void TestIsCellInsideRange()
        {
            var wb = new XLWorkbook();
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

        [Test]
        public void TestIsRangeInsideAnotherRange()
        {
            var wb = new XLWorkbook();
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

        [Test]
        public void TestGetNearestParentRange()
        {
            var wb = new XLWorkbook();
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
            ExceptionAssert.Throws<InvalidOperationException>(() => ExcelHelper.GetNearestParentRange(new[] { parentRange1, parentRange2, notParentRange, parentRange3, parentRange4 }, childRange),
                "Found more than one nearest parent ranges");

            IXLRange badChildRange = ws.Range(7, 7, 7, 8);
            ExceptionAssert.Throws<InvalidOperationException>(() => ExcelHelper.GetNearestParentRange(new[] { parentRange1, parentRange2, notParentRange, parentRange3, parentRange4 }, badChildRange),
                "Nearest parent range was not found");
        }

        [Test]
        public void TestGetCellCoordsRelativeRange()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLRange range = ws.Range(3, 3, 6, 7);
            Assert.AreEqual(new CellCoords(1, 1), ExcelHelper.GetCellCoordsRelativeRange(range, ws.Cell(3, 3)));
            Assert.AreEqual(new CellCoords(1, 2), ExcelHelper.GetCellCoordsRelativeRange(range, ws.Cell(3, 4)));
            Assert.AreEqual(new CellCoords(1, 5), ExcelHelper.GetCellCoordsRelativeRange(range, ws.Cell(3, 7)));
            Assert.AreEqual(new CellCoords(2, 1), ExcelHelper.GetCellCoordsRelativeRange(range, ws.Cell(4, 3)));
            Assert.AreEqual(new CellCoords(2, 4), ExcelHelper.GetCellCoordsRelativeRange(range, ws.Cell(4, 6)));
            Assert.AreEqual(new CellCoords(4, 3), ExcelHelper.GetCellCoordsRelativeRange(range, ws.Cell(6, 5)));
            Assert.AreEqual(new CellCoords(4, 5), ExcelHelper.GetCellCoordsRelativeRange(range, ws.Cell(6, 7)));

            IXLCell cell1 = ws.Cell(1, 1);
            IXLCell cell2 = ws.Cell(7, 8);
            ExceptionAssert.Throws<InvalidOperationException>(() => ExcelHelper.GetCellCoordsRelativeRange(range, cell1), $"Cell {cell1} is outside of the range {range}");
            ExceptionAssert.Throws<InvalidOperationException>(() => ExcelHelper.GetCellCoordsRelativeRange(range, cell2), $"Cell {cell2} is outside of the range {range}");

            Assert.AreEqual(new CellCoords(-1, -1), ExcelHelper.GetCellCoordsRelativeRange(range, cell1, false));
            Assert.AreEqual(new CellCoords(5, 6), ExcelHelper.GetCellCoordsRelativeRange(range, cell2, false));
            Assert.AreEqual(new CellCoords(0, 0), ExcelHelper.GetCellCoordsRelativeRange(range, ws.Cell(2, 2), false));
            Assert.AreEqual(new CellCoords(0, 2), ExcelHelper.GetCellCoordsRelativeRange(range, ws.Cell(2, 4), false));
            Assert.AreEqual(new CellCoords(2, 0), ExcelHelper.GetCellCoordsRelativeRange(range, ws.Cell(4, 2), false));
        }

        [Test]
        public void TestGetRangeCoordsRelativeParent()
        {
            var wb = new XLWorkbook();
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
            ExceptionAssert.Throws<InvalidOperationException>(() => ExcelHelper.GetRangeCoordsRelativeParent(parentRange, childRange), $"Range {parentRange} is not a parent of the range {childRange}. Child range is outside of the parent range.");

            Assert.AreEqual(new RangeCoords(new CellCoords(2, 2), new CellCoords(5, 3)), ExcelHelper.GetRangeCoordsRelativeParent(parentRange, childRange, false));

            childRange = ws.Range(1, 1, 2, 1);
            Assert.AreEqual(new RangeCoords(new CellCoords(-1, -1), new CellCoords(0, -1)), ExcelHelper.GetRangeCoordsRelativeParent(parentRange, childRange, false));

            childRange = ws.Range(1, 1, 3, 3);
            Assert.AreEqual(new RangeCoords(new CellCoords(-1, -1), new CellCoords(1, 1)), ExcelHelper.GetRangeCoordsRelativeParent(parentRange, childRange, false));

            childRange = ws.Range(8, 6, 9, 7);
            Assert.AreEqual(new RangeCoords(new CellCoords(6, 4), new CellCoords(7, 5)), ExcelHelper.GetRangeCoordsRelativeParent(parentRange, childRange, false));
        }

        [Test]
        public void TestGetAddressShift()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLCell cell1 = ws.Cell(1, 1);
            IXLCell cell2 = ws.Cell(5, 4);
            Assert.AreEqual(new AddressShift(4, 3), ExcelHelper.GetAddressShift(cell2.Address, cell1.Address));
            Assert.AreEqual(new AddressShift(-4, -3), ExcelHelper.GetAddressShift(cell1.Address, cell2.Address));
        }

        [Test]
        public void TestShiftCell()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLCell cell = ws.Cell(1, 1);
            var shift = new AddressShift(4, 3);

            Assert.AreEqual(ws.Cell(5, 4), ExcelHelper.ShiftCell(cell, shift));

            cell = ws.Cell(5, 4);
            shift = new AddressShift(-4, -3);
            Assert.AreEqual(ws.Cell(1, 1), ExcelHelper.ShiftCell(cell, shift));
        }

        [Test]
        public void TestCopyRange()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");
            IXLRange range = ws.Range(6, 5, 9, 7);

            range.FirstCell().Value = "RangeStart";
            range.LastCell().Value = "RangeEnd";
            range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.LastCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);

            ws.Cell(11, 9).Value = DateTime.Now;

            IXLRange newRange = ExcelHelper.CopyRange(range, ws.Cell(10, 8));

            Assert.AreEqual(4, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
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
            Assert.AreEqual(XLDataType.Blank, ws.Cell(11, 9).DataType);
            Assert.AreEqual(string.Empty, ws.Cell(11, 9).Value.ToString());

            range.Clear();
            IXLRange newRange2 = ExcelHelper.CopyRange(newRange, ws.Cell(11, 8));

            Assert.AreEqual(3, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
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

        [Test]
        public void TestCopyNamedRange()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");
            IXLRange range = ws.Range(6, 5, 9, 7);
            range.AddToNamed("TestRange", XLScope.Worksheet);
            IXLNamedRange namedRange = ws.NamedRange("TestRange");

            range.FirstCell().Value = "RangeStart";
            range.LastCell().Value = "RangeEnd";
            range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.LastCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);

            ws.Cell(11, 9).Value = DateTime.Now;

            IXLNamedRange copiedNamedRange = ExcelHelper.CopyNamedRange(namedRange, ws.Cell(10, 8), "Copy");
            IXLRange newRange = copiedNamedRange.Ranges.ElementAt(0);

            Assert.AreEqual(1, copiedNamedRange.Ranges.Count);
            Assert.AreEqual(4, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
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
            Assert.AreEqual(XLDataType.Blank, ws.Cell(11, 9).DataType);
            Assert.AreEqual(string.Empty, ws.Cell(11, 9).Value.ToString());

            range.Clear();
            IXLNamedRange copiedNamedRange2 = ExcelHelper.CopyNamedRange(copiedNamedRange, ws.Cell(11, 8), "Copy2");
            IXLRange newRange2 = copiedNamedRange2.Ranges.ElementAt(0);

            Assert.AreEqual(1, copiedNamedRange2.Ranges.Count);
            Assert.AreEqual(3, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
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

        [Test]
        public void TestAllocateSpaceForNextRange()
        {
            // Adding cells at the top
            var wb = InitWorkBookForShiftTests();
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

            Assert.AreEqual(10, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
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


            // Adding cells at the bottom
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

            Assert.AreEqual(10, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
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


            // Adding rows at the top
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

            Assert.AreEqual(10, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
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


            // Adding rows at the bottom
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

            Assert.AreEqual(10, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
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


            // Adding cells on the left
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

            Assert.AreEqual(10, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
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


            // Adding cells on the right
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

            Assert.AreEqual(10, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
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


            // Adding columns on the left
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

            Assert.AreEqual(10, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
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


            // Adding columns on the right
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

            Assert.AreEqual(10, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
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


            // Adding nothing (without shift)
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

            Assert.AreEqual(10, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
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

        [Test]
        public void TestDeleteRange()
        {
            // Deleting with moving cells up
            var wb = InitWorkBookForShiftTests();
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
            Assert.AreEqual(8, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
            Assert.AreEqual(belowCell1, ws.Cell(6, 6));
            Assert.AreEqual(belowCell2, ws.Cell(10, 8));
            Assert.AreEqual(rightCell1, ws.Cell(7, 8));
            Assert.AreEqual(rightCell2, ws.Cell(5, 8));
            Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
            Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
            Assert.AreEqual(leftCell1, ws.Cell(7, 4));
            Assert.AreEqual(leftCell2, ws.Cell(10, 4));

            // Deleting with moving the row up
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
            Assert.AreEqual(6, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
            Assert.AreEqual(belowCell1, ws.Cell(6, 6));
            Assert.AreEqual(belowCell2, ws.Cell(6, 8));
            Assert.AreEqual(rightCell2, ws.Cell(5, 8));
            Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
            Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
            Assert.AreEqual(leftCell2, ws.Cell(6, 4));

            // Deleting with moving cells left
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
            Assert.AreEqual(8, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
            Assert.AreEqual(belowCell1, ws.Cell(10, 6));
            Assert.AreEqual(belowCell2, ws.Cell(10, 8));
            Assert.AreEqual(rightCell1, ws.Cell(7, 5));
            Assert.AreEqual(rightCell2, ws.Cell(5, 8));
            Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
            Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
            Assert.AreEqual(leftCell1, ws.Cell(7, 4));
            Assert.AreEqual(leftCell2, ws.Cell(10, 4));

            // Deleting with moving the column left
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
            Assert.AreEqual(6, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
            Assert.AreEqual(belowCell2, ws.Cell(10, 5));
            Assert.AreEqual(rightCell1, ws.Cell(7, 5));
            Assert.AreEqual(rightCell2, ws.Cell(5, 5));
            Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
            Assert.AreEqual(leftCell1, ws.Cell(7, 4));
            Assert.AreEqual(leftCell2, ws.Cell(10, 4));

            // Deleting without any shift
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
            Assert.AreEqual(8, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
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

        [Test]
        public void TestMoveRange()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");
            IXLRange range = ws.Range(6, 5, 9, 7);

            range.FirstCell().Value = "RangeStart";
            range.LastCell().Value = "RangeEnd";
            range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.LastCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);

            ws.Cell(11, 9).Value = DateTime.Now;

            IXLRange movedRange = ExcelHelper.MoveRange(range, ws.Cell(10, 8));

            Assert.AreEqual(2, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
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
            Assert.AreEqual(XLDataType.Blank, ws.Cell(11, 9).DataType);
            Assert.AreEqual(string.Empty, ws.Cell(11, 9).Value.ToString());

            IXLRange movedRange2 = ExcelHelper.MoveRange(movedRange, ws.Cell(11, 8));

            Assert.AreEqual(2, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
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

        [Test]
        public void TestMoveNamedRange()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");
            IXLRange range = ws.Range(6, 5, 9, 7);
            range.AddToNamed("TestRange", XLScope.Worksheet);
            IXLNamedRange namedRange = ws.NamedRange("TestRange");

            range.FirstCell().Value = "RangeStart";
            range.LastCell().Value = "RangeEnd";
            range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.LastCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);

            ws.Cell(11, 9).Value = DateTime.Now;

            IXLNamedRange movedNamedRange = ExcelHelper.MoveNamedRange(namedRange, ws.Cell(10, 8));
            IXLRange movedRange = movedNamedRange.Ranges.ElementAt(0);

            Assert.AreEqual(movedNamedRange, ws.NamedRange("TestRange"));
            Assert.AreEqual(1, movedNamedRange.Ranges.Count);
            Assert.AreEqual(2, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
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
            Assert.AreEqual(XLDataType.Blank, ws.Cell(11, 9).DataType);
            Assert.AreEqual(string.Empty, ws.Cell(11, 9).Value.ToString());

            IXLNamedRange movedNamedRange2 = ExcelHelper.MoveNamedRange(movedNamedRange, ws.Cell(11, 8));
            IXLRange movedRange2 = movedNamedRange2.Ranges.ElementAt(0);

            Assert.AreEqual(movedNamedRange2, ws.NamedRange("TestRange"));
            Assert.AreEqual(1, movedNamedRange2.Ranges.Count);
            Assert.AreEqual(2, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
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
            var wb = new XLWorkbook();
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

        [Test]
        public void TestAddTempWorksheet()
        {
            var wb = new XLWorkbook();
            ExcelHelper.AddTempWorksheet(wb);

            Assert.AreEqual(1, wb.Worksheets.Count);
            Assert.IsTrue(Regex.IsMatch(wb.Worksheets.First().Name, "[0-9a-f]{31}"));
        }

        [Test]
        public void TestMergeRanges()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");
            IXLWorksheet ws2 = wb.AddWorksheet("Test2");
            IXLRange range1 = ws.Range(3, 3, 5, 5);
            IXLRange range2 = ws.Range(1, 7, 2, 8);
            IXLRange range3 = ws2.Range(1, 7, 2, 8);

            Assert.AreEqual(ws.Range(ws.Cell(1, 3), ws.Cell(5, 8)), ExcelHelper.MergeRanges(range1, range2));

            range2 = ws.Range(6, 7, 7, 8);
            Assert.AreEqual(ws.Range(ws.Cell(3, 3), ws.Cell(7, 8)), ExcelHelper.MergeRanges(range1, range2));

            range2 = ws.Range(4, 7, 4, 8);
            Assert.AreEqual(ws.Range(ws.Cell(3, 3), ws.Cell(5, 8)), ExcelHelper.MergeRanges(range1, range2));

            range2 = ws.Range(4, 1, 4, 2);
            Assert.AreEqual(ws.Range(ws.Cell(3, 1), ws.Cell(5, 5)), ExcelHelper.MergeRanges(range1, range2));

            range2 = ws.Range(4, 4, 4, 4);
            Assert.AreEqual(range1, ExcelHelper.MergeRanges(range1, range2));

            range2 = ws.Range(2, 2, 3, 4);
            Assert.AreEqual(ws.Range(ws.Cell(2, 2), ws.Cell(5, 5)), ExcelHelper.MergeRanges(range1, range2));

            range2 = ws.Range(4, 4, 6, 6);
            Assert.AreEqual(ws.Range(ws.Cell(3, 3), ws.Cell(6, 6)), ExcelHelper.MergeRanges(range1, range2));

            Assert.AreEqual(range1, ExcelHelper.MergeRanges(range1, null));
            Assert.AreEqual(range2, ExcelHelper.MergeRanges(null, range2));
            Assert.IsNull(ExcelHelper.MergeRanges(null, null));

            ExceptionAssert.Throws<InvalidOperationException>(() => ExcelHelper.MergeRanges(range1, range3), "Ranges belong to different worksheets");

            range1.Delete(XLShiftDeletedCells.ShiftCellsLeft);

            Assert.AreEqual(range2, ExcelHelper.MergeRanges(range1, range2));
            Assert.IsNull(ExcelHelper.MergeRanges(range1, null));

            range1 = ws.Range(3, 3, 5, 5);
            range2.Delete(XLShiftDeletedCells.ShiftCellsLeft);

            Assert.AreEqual(range1, ExcelHelper.MergeRanges(range1, range2));
            Assert.IsNull(ExcelHelper.MergeRanges(null, range2));

            range1.Delete(XLShiftDeletedCells.ShiftCellsLeft);

            Assert.IsNull(ExcelHelper.MergeRanges(range1, range2));
        }

        [Test]
        public void TestIsRangeInvalid()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            var range = ws.Range(3, 3, 5, 5);
            Assert.IsFalse(ExcelHelper.IsRangeInvalid(range));

            range.Delete(XLShiftDeletedCells.ShiftCellsUp);
            Assert.IsTrue(ExcelHelper.IsRangeInvalid(range));

            Assert.IsFalse(ExcelHelper.IsRangeInvalid(ws.Range(2, 2, 2, 2)));
            Assert.IsFalse(ExcelHelper.IsRangeInvalid(ws.Range(1, 2, 2, 3)));
            Assert.IsTrue(ExcelHelper.IsRangeInvalid(ws.Range(2, 2, 1, 3)));
            Assert.IsTrue(ExcelHelper.IsRangeInvalid(ws.Range(1, 3, 2, 2)));
            Assert.IsTrue(ExcelHelper.IsRangeInvalid(ws.Range(2, 3, 1, 2)));
        }

        [Test]
        public void TestGetMaxCell()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLRange range = ws.Range(3, 3, 5, 5);

            Assert.AreEqual(ws.Cell(5, 5), ExcelHelper.GetMaxCell(range.Cells().Concat(new[] { ws.Cell(2, 1) }).ToArray()));
            Assert.AreEqual(ws.Cell(5, 5), ExcelHelper.GetMaxCell(range.Cells().Concat(new[] { ws.Cell(5, 1) }).ToArray()));
            Assert.AreEqual(ws.Cell(5, 5), ExcelHelper.GetMaxCell(range.Cells().Concat(new[] { ws.Cell(1, 3) }).ToArray()));
            Assert.AreEqual(ws.Cell(5, 5), ExcelHelper.GetMaxCell(range.Cells().Concat(new[] { ws.Cell(1, 5) }).ToArray()));
            Assert.AreEqual(ws.Cell(5, 5), ExcelHelper.GetMaxCell(range.Cells().Concat(new[] { ws.Cell(4, 4) }).ToArray()));
            Assert.AreEqual(ws.Cell(5, 5), ExcelHelper.GetMaxCell(range.Cells().Concat(new[] { ws.Cell(5, 5) }).ToArray()));
            Assert.AreEqual(ws.Cell(6, 5), ExcelHelper.GetMaxCell(range.Cells().Concat(new[] { ws.Cell(6, 1) }).ToArray()));
            Assert.AreEqual(ws.Cell(10, 5), ExcelHelper.GetMaxCell(range.Cells().Concat(new[] { ws.Cell(10, 4) }).ToArray()));
            Assert.AreEqual(ws.Cell(10, 10), ExcelHelper.GetMaxCell(range.Cells().Concat(new[] { ws.Cell(10, 10) }).ToArray()));
            Assert.AreEqual(ws.Cell(5, 6), ExcelHelper.GetMaxCell(range.Cells().Concat(new[] { ws.Cell(1, 6) }).ToArray()));
            Assert.AreEqual(ws.Cell(5, 10), ExcelHelper.GetMaxCell(range.Cells().Concat(new[] { ws.Cell(3, 10) }).ToArray()));
            Assert.AreEqual(ws.Cell(20, 20), ExcelHelper.GetMaxCell(new[] { ws.Cell(20, 20) }));

            Assert.IsNull(ExcelHelper.GetMaxCell(null));
            Assert.IsNull(ExcelHelper.GetMaxCell(Enumerable.Empty<IXLCell>().ToArray()));
        }

        [Test]
        public void TestCloneRange()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLRange range = ws.Range(3, 3, 5, 5);

            Assert.AreEqual(range, ExcelHelper.CloneRange(range));

            //// Since ClosedXml 0.93.0 this test does not pass
            //Assert.AreNotSame(range, ExcelHelper.CloneRange(range));

            Assert.IsNull(ExcelHelper.CloneRange(null));
        }
    }
}