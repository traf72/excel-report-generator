using ClosedXML.Excel;

namespace ExcelReportGenerator.Extensions;

internal static class XLRangeBaseExtensions
{
    public static IXLCells CellsUsedWithoutFormulas(this IXLRangeBase range)
    {
        return range.CellsUsed(c => !c.HasFormula);
    }

    public static IXLCells CellsUsedWithoutFormulas(this IXLRangeBase range, Func<IXLCell, bool> predicate)
    {
        return range.CellsUsed(c => predicate(c) && !c.HasFormula);
    }

    public static IXLCells CellsUsedWithoutFormulas(this IXLRangeBase range, XLCellsUsedOptions options)
    {
        return range.CellsUsed(options, c => !c.HasFormula);
    }

    public static IXLCells CellsUsedWithoutFormulas(this IXLRangeBase range, XLCellsUsedOptions options, Func<IXLCell, bool> predicate)
    {
        return range.CellsUsed(options, c => predicate(c) && !c.HasFormula);
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