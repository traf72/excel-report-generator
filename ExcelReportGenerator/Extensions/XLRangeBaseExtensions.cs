using ClosedXML.Excel;
using System;

namespace ExcelReportGenerator.Extensions
{
    public static class XLRangeBaseExtensions
    {
        public static IXLCells CellsUsedWithoutFormulas(this IXLRangeBase range)
        {
            return range.CellsUsed(c => !c.HasFormula);
        }

        public static IXLCells CellsUsedWithoutFormulas(this IXLRangeBase range, Func<IXLCell, bool> predicate)
        {
            return range.CellsUsed(c => predicate(c) && !c.HasFormula);
        }

        public static IXLCells CellsUsedWithoutFormulas(this IXLRangeBase range, bool includeFormats)
        {
            return range.CellsUsed(includeFormats, c => !c.HasFormula);
        }

        public static IXLCells CellsUsedWithoutFormulas(this IXLRangeBase range, bool includeFormats, Func<IXLCell, bool> predicate)
        {
            return range.CellsUsed(includeFormats, c => predicate(c) && !c.HasFormula);
        }

        public static IXLCells CellsWithoutFormulas(this IXLRangeBase range)
        {
            return range.Cells(c => !c.HasFormula);
        }

        public static IXLCells CellsWithoutFormulas(this IXLRangeBase range, Func<IXLCell, bool> predicate)
        {
            return range.Cells(c => predicate(c) && !c.HasFormula);
        }
    }
}