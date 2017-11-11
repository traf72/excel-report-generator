using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ExcelReporter.Enums;

namespace ExcelReporter.Excel
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
            IEnumerable<IXLRange> nearestParents = parents.Where(p => p.Cells().Count() == cellsCountInMinRange).ToList();
            if (nearestParents.Count() > 1)
            {
                throw new InvalidOperationException("Found more than one nearest parent ranges");
            }

            return nearestParents.First();
        }

        public static CellCoords GetCellCoordsRelativeRange(IXLRange range, IXLCell cell)
        {
            if (!IsCellInsideRange(cell, range))
            {
                throw new InvalidOperationException($"Cell {cell} is outside of the range {range}");
            }

            return new CellCoords(cell.Address.RowNumber - range.FirstRow().RowNumber() + 1,
                cell.Address.ColumnNumber - range.FirstColumn().ColumnNumber() + 1);
        }

        public static RangeCoords GetRangeCoordsRelativeParent(IXLRange parentRange, IXLRange childRange)
        {
            if (!IsRangeInsideAnotherRange(parentRange, childRange))
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
                // Если ячейка, в которую производится копирование, находится внутри диапазона, то копирование происходит неверно
                // (копирование происходит по ячейкам, скопированные ячейки сразу же появляются в первом диапазоне и начинают копироваться поновой)
                // Поэтому копируем через вспомогательный лист
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

            // Если ячейка, в которую производится перемещение, находится внутри диапазона, то верхний способ не подойдёт
            // Поэтому перемещаем через вспомогательный лист
            IXLWorksheet tempWs = null;
            try
            {
                tempWs = AddTempWorksheet(range.Worksheet.Workbook);
                IXLRange tempRange = CopyRange(range, tempWs.FirstCell());
                range.Clear();
                return CopyRange(tempRange, cell);
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

        public static IXLWorksheet AddTempWorksheet(XLWorkbook wb)
        {
            // Отсекаем один символ от Guid'а, так как наименование листа не может быть больше 31 символа
            return wb.AddWorksheet(Guid.NewGuid().ToString("N").Substring(1));
        }
    }
}