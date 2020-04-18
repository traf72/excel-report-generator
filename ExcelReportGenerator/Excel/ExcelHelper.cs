using ClosedXML.Excel;
using ExcelReportGenerator.Enums;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReportGenerator.Excel
{
    internal static class ExcelHelper
    {
        public static bool IsCellInsideRange(IXLCell cell, IXLRange range)
        {
            return range.Cells().Contains(cell);
        }

        public static bool IsRangeInsideAnotherRange(IXLRange parentRange, IXLRange childRange)
        {
            return parentRange.Worksheet == childRange.Worksheet
                && childRange.FirstRow().RowNumber() >= parentRange.FirstRow().RowNumber()
                && childRange.LastRow().RowNumber() <= parentRange.LastRow().RowNumber()
                && childRange.FirstColumn().ColumnNumber() >= parentRange.FirstColumn().ColumnNumber()
                && childRange.LastColumn().ColumnNumber() <= parentRange.LastColumn().ColumnNumber();
        }

        public static IXLRange GetNearestParentRange(IEnumerable<IXLRange> parentRanges, IXLRange range)
        {
            IList<IXLRange> parents = new List<IXLRange>();
            foreach (IXLRange parent in parentRanges)
            {
                if (!IsRangeInsideAnotherRange(parent, range))
                {
                    continue;
                }
                parents.Add(parent);
            }

            if (parents.Count == 1)
            {
                return parents.First();
            }
            if (!parents.Any())
            {
                throw new InvalidOperationException("Nearest parent range was not found");
            }

            int cellsCountInMinRange = parents.Min(p => p.Cells().Count());
            IList<IXLRange> nearestParents = parents.Where(p => p.Cells().Count() == cellsCountInMinRange).ToList();
            if (nearestParents.Count > 1)
            {
                throw new InvalidOperationException("Found more than one nearest parent ranges");
            }

            return nearestParents.First();
        }

        public static CellCoords GetCellCoordsRelativeRange(IXLRange range, IXLCell cell, bool checkСorrectness = true)
        {
            if (checkСorrectness && !IsCellInsideRange(cell, range))
            {
                throw new InvalidOperationException($"Cell {cell} is outside of the range {range}");
            }

            return new CellCoords(cell.Address.RowNumber - range.FirstRow().RowNumber() + 1,
                cell.Address.ColumnNumber - range.FirstColumn().ColumnNumber() + 1);
        }

        public static RangeCoords GetRangeCoordsRelativeParent(IXLRange parentRange, IXLRange childRange, bool checkСorrectness = true)
        {
            if (checkСorrectness && !IsRangeInsideAnotherRange(parentRange, childRange))
            {
                throw new InvalidOperationException($"Range {parentRange} is not a parent of the range {childRange}. Child range is outside of the parent range.");
            }

            CellCoords firstCell = new CellCoords(childRange.FirstRow().RowNumber() - parentRange.FirstRow().RowNumber() + 1,
                childRange.FirstColumn().ColumnNumber() - parentRange.FirstColumn().ColumnNumber() + 1);

            CellCoords lastCell = new CellCoords(childRange.LastRow().RowNumber() - parentRange.FirstRow().RowNumber() + 1,
                childRange.LastColumn().ColumnNumber() - parentRange.FirstColumn().ColumnNumber() + 1);

            return new RangeCoords(firstCell, lastCell);
        }

        public static IXLNamedRange CopyNamedRange(IXLNamedRange namedRange, IXLCell cell, string name)
        {
            IXLRange newRange = CopyRange(namedRange.Ranges.ElementAt(0), cell);
            newRange.AddToNamed(name, XLScope.Worksheet);
            return cell.Worksheet.NamedRange(name);
        }

        public static IXLRange CopyRange(IXLRange range, IXLCell cell)
        {
            IXLCell newRangeFirstCell = cell;
            IXLCell newRangeLastCell = ShiftCell(newRangeFirstCell, new AddressShift(range.RowCount() - 1, range.ColumnCount() - 1));
            IXLRange newRange = range.Worksheet.Range(newRangeFirstCell, newRangeLastCell);
            if (!IsCellInsideRange(cell, range))
            {
                newRange.Clear();
                range.CopyTo(newRange);
            }
            else
            {
                // If the cell which the copy occurs to is inside the range then the copy will be wrong (copying is performed by cells,
                // copied cells appear in the first range immediately and start copying again)
                // That's why, copy through the auxiliary sheet
                IXLWorksheet tempWs = null;
                try
                {
                    tempWs = AddTempWorksheet(range.Worksheet.Workbook);
                    IXLRange tempRange = range.CopyTo(tempWs.FirstCell());
                    newRange.Clear();
                    tempRange.CopyTo(newRange);
                }
                finally
                {
                    tempWs?.Delete();
                }
            }

            return newRange;
        }

        public static AddressShift GetAddressShift(IXLAddress address1, IXLAddress address2)
        {
            return new AddressShift(address1.RowNumber - address2.RowNumber, address1.ColumnNumber - address2.ColumnNumber);
        }

        public static IXLCell ShiftCell(IXLCell cell, AddressShift shift)
        {
            return cell.Worksheet.Cell(cell.Address.RowNumber + shift.RowCount, cell.Address.ColumnNumber + shift.ColCount);
        }

        public static void AllocateSpaceForNextRange(IXLRange copiedRange, Direction direction, ShiftType type = ShiftType.Cells)
        {
            switch (direction)
            {
                case Direction.Bottom:
                    if (type == ShiftType.Cells)
                    {
                        copiedRange.InsertRowsBelow(copiedRange.RowCount(), false);
                    }
                    else if (type == ShiftType.Row)
                    {
                        copiedRange.Worksheet.Row(copiedRange.LastRow().RowNumber()).InsertRowsBelow(copiedRange.RowCount());
                    }
                    break;

                case Direction.Right:
                    if (type == ShiftType.Cells)
                    {
                        copiedRange.InsertColumnsAfter(copiedRange.ColumnCount(), false);
                    }
                    else if (type == ShiftType.Row)
                    {
                        copiedRange.Worksheet.Column(copiedRange.LastColumn().ColumnNumber()).InsertColumnsAfter(copiedRange.ColumnCount());
                    }
                    break;

                case Direction.Top:
                    if (type == ShiftType.Cells)
                    {
                        copiedRange.InsertRowsAbove(copiedRange.RowCount(), false);
                    }
                    else if (type == ShiftType.Row)
                    {
                        copiedRange.Worksheet.Row(copiedRange.FirstRow().RowNumber()).InsertRowsAbove(copiedRange.RowCount());
                    }
                    break;

                case Direction.Left:
                    if (type == ShiftType.Cells)
                    {
                        copiedRange.InsertColumnsBefore(copiedRange.ColumnCount(), false);
                    }
                    else if (type == ShiftType.Row)
                    {
                        copiedRange.Worksheet.Column(copiedRange.FirstColumn().ColumnNumber()).InsertColumnsBefore(copiedRange.ColumnCount());
                    }
                    break;
            }
        }

        public static void DeleteRange(IXLRange range, ShiftType type, XLShiftDeletedCells shiftDirection = XLShiftDeletedCells.ShiftCellsUp)
        {
            switch (type)
            {
                case ShiftType.Cells:
                    range.Delete(shiftDirection);
                    break;

                case ShiftType.Row:
                    if (shiftDirection == XLShiftDeletedCells.ShiftCellsUp)
                    {
                        range.Worksheet.Rows(range.FirstRow().RowNumber(), range.LastRow().RowNumber()).Delete();
                    }
                    else
                    {
                        range.Worksheet.Columns(range.FirstColumn().ColumnNumber(), range.LastColumn().ColumnNumber()).Delete();
                    }
                    break;

                case ShiftType.NoShift:
                    range.Clear();
                    break;
            }
        }

        public static IXLRange MoveRange(IXLRange range, IXLCell cell)
        {
            if (!IsCellInsideRange(cell, range))
            {
                IXLRange newRange = CopyRange(range, cell);
                range.Clear();
                return newRange;
            }


            // If the cell which the movement occurs to is inside the range then the way above will not work properly,
            // That's why, copy through the auxiliary sheet
            IXLWorksheet tempWs = null;
            try
            {
                tempWs = AddTempWorksheet(range.Worksheet.Workbook);
                IXLRange tempRange = range.CopyTo(tempWs.FirstCell());
                range.Clear();
                return tempRange.CopyTo(cell);
            }
            finally
            {
                tempWs?.Delete();
            }
        }

        public static IXLNamedRange MoveNamedRange(IXLNamedRange namedRange, IXLCell cell)
        {
            string name = namedRange.Name;
            IXLRange newRange = MoveRange(namedRange.Ranges.ElementAt(0), cell);
            namedRange.Delete();
            newRange.AddToNamed(name, XLScope.Worksheet);
            return cell.Worksheet.NamedRange(name);
        }

        public static IXLRange MergeRanges(IXLRange range1, IXLRange range2)
        {
            if (range1 == null || IsRangeInvalid(range1))
            {
                return range2 == null || IsRangeInvalid(range2) ? null : range2;
            }

            if (range2 == null || IsRangeInvalid(range2))
            {
                return IsRangeInvalid(range1) ? null : range1;
            }

            if (range1.Worksheet != range2.Worksheet)
            {
                throw new InvalidOperationException("Ranges belong to different worksheets");
            }

            IXLWorksheet ws = range1.Worksheet;

            IXLCell newRangeFirstCell = ws.Cell(Math.Min(range1.FirstRow().RowNumber(), range2.FirstRow().RowNumber()),
                Math.Min(range1.FirstColumn().ColumnNumber(), range2.FirstColumn().ColumnNumber()));
            IXLCell newRangeLastCell = ws.Cell(Math.Max(range1.LastRow().RowNumber(), range2.LastRow().RowNumber()),
                Math.Max(range1.LastColumn().ColumnNumber(), range2.LastColumn().ColumnNumber()));

            return ws.Range(newRangeFirstCell, newRangeLastCell);
        }

        public static bool IsRangeInvalid(IXLRange range)
        {
            if (!range.RangeAddress.IsValid)
            {
                return true;
            }

            try
            {
                range.FirstRow().RowNumber();
                range.LastRow().RowNumber();
                range.FirstColumn().ColumnNumber();
                range.LastColumn().ColumnNumber();
            }
            catch (ArgumentOutOfRangeException)
            {
                return true;
            }

            return false;
        }

        public static IXLCell GetMaxCell(IXLCell[] cells)
        {
            if (cells == null || !cells.Any())
            {
                return null;
            }

            IXLCell cellWithMaxRowNum = cells.First(c1 => c1.Address.RowNumber == cells.Max(c2 => c2.Address.RowNumber));
            IXLCell cellWithMaxColumnNum = cells.First(c1 => c1.Address.ColumnNumber == cells.Max(c2 => c2.Address.ColumnNumber));

            IXLWorksheet ws = cellWithMaxRowNum.Worksheet;
            IXLCell maxCell = ws.Cell(
                Math.Max(cellWithMaxRowNum.Address.RowNumber, cellWithMaxColumnNum.Address.RowNumber),
                Math.Max(cellWithMaxRowNum.Address.ColumnNumber, cellWithMaxColumnNum.Address.ColumnNumber));

            return maxCell;
        }

        public static IXLRange CloneRange(IXLRange range)
        {
            if (range == null)
            {
                return null;
            }

            IXLAddress firstCellAddress = range.FirstCell().Address;
            IXLAddress lastCellAddress = range.LastCell().Address;

            return range.Worksheet.Range(firstCellAddress.RowNumber, firstCellAddress.ColumnNumber,
                lastCellAddress.RowNumber, lastCellAddress.ColumnNumber);
        }

        public static IXLWorksheet AddTempWorksheet(XLWorkbook wb)
        {
            // Trim one character from Guid because the name of a sheet can't be more than 31 symbols
            return wb.AddWorksheet(Guid.NewGuid().ToString("N").Substring(1));
        }
    }
}