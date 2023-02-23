using ClosedXML.Excel;
using NUnit.Framework;
using System.Linq;

namespace ExcelReportGenerator.Tests.CustomAsserts
{
    public class ExcelAssert
    {
        public static void AreWorksheetsContentEquals(IXLWorksheet expected, IXLWorksheet actual)
        {
            if (expected == actual)
            {
                return;
            }

            Assert.AreEqual(expected.CellsUsed(XLCellsUsedOptions.All).Count(), actual.CellsUsed(XLCellsUsedOptions.All).Count(), "Cells used count failed");

            IXLCell expectedFirstCellUsed = expected.FirstCellUsed(XLCellsUsedOptions.All);
            IXLCell actualFirstCellUsed = actual.FirstCellUsed(XLCellsUsedOptions.All);
            Assert.AreEqual(expectedFirstCellUsed.Address, actualFirstCellUsed.Address, "First cell used failed");
            IXLCell expectedLastCellUsed = expected.LastCellUsed(XLCellsUsedOptions.All);
            IXLCell actualLastCellUsed = actual.LastCellUsed(XLCellsUsedOptions.All);
            Assert.AreEqual(expectedLastCellUsed.Address, actualLastCellUsed.Address, "Last cell used failed");

            IXLRange range = expected.Range(expectedFirstCellUsed, expectedLastCellUsed);
            foreach (IXLCell expectedCell in range.Cells())
            {
                IXLCell actualCell = actual.Cell(expectedCell.Address);
                if (expectedCell.HasFormula)
                {
                    Assert.AreEqual(expectedCell.FormulaA1, actualCell.FormulaA1, $"Cell {expectedCell.Address} FormulaA1 failed.");
                    Assert.AreEqual(expectedCell.FormulaR1C1, actualCell.FormulaR1C1, $"Cell {expectedCell.Address} FormulaR1C1 failed.");
                    Assert.AreEqual(expectedCell.FormulaReference, actualCell.FormulaReference, $"Cell {expectedCell.Address} FormulaReference failed.");
                }
                else
                {
                    Assert.AreEqual(expectedCell.Value, actualCell.Value, $"Cell {expectedCell.Address} Value failed.");
                }
                Assert.AreEqual(expectedCell.DataType, actualCell.DataType, $"Cell {expectedCell.Address} DataType failed.");
                Assert.AreEqual(expectedCell.Active, actualCell.Active, $"Cell {expectedCell.Address} Active failed.");
                AreColumnsEquals(expectedCell.WorksheetColumn(), actualCell.WorksheetColumn(), $"Column {actualCell.WorksheetColumn().RangeAddress} {{0}} failed.");
                AreRowEquals(expectedCell.WorksheetRow(), actualCell.WorksheetRow(), $"Row {actualCell.WorksheetRow().RangeAddress} {{0}} failed.");
                AreCellsStyleEquals(expectedCell.Style, actualCell.Style, $"Cell {expectedCell.Address} Style {{0}} failed.");
                AreCellsCommentEquals(expectedCell.GetComment(), actualCell.GetComment(), $"Cell {expectedCell.Address} Comment {{0}} failed.");
            }

            AreMergedRangesEquals(expected.MergedRanges, actual.MergedRanges);
            AreNamedRangesEquals(expected.NamedRanges, actual.NamedRanges);
            ArePageSetupEquals(expected.PageSetup, actual.PageSetup, "PageSetup {0} failed.");
        }

        public static void AreWorkbooksContentEquals(XLWorkbook expected, XLWorkbook actual)
        {
            if (expected == actual)
            {
                return;
            }

            Assert.AreEqual(expected.Worksheets.Count, actual.Worksheets.Count, "Workbook worksheets count failed");
            for (int i = 0; i < expected.Worksheets.Count; i++)
            {
                Assert.AreEqual(expected.Worksheet(i + 1).Name, actual.Worksheet(i + 1).Name, "Worksheets names failed");
            }

            Assert.AreEqual(expected.NamedRanges.Count(), actual.NamedRanges.Count(), "Workbook named ranges count failed");
            AreNamedRangesEquals(expected.NamedRanges, actual.NamedRanges);
            for (int i = 0; i < expected.Worksheets.Count; i++)
            {
                AreWorksheetsContentEquals(expected.Worksheet(i + 1), actual.Worksheet(i + 1));
            }
        }

        public static void AreColumnsEquals(IXLColumn expected, IXLColumn actual, string message = null)
        {
            if (expected.Equals(actual))
            {
                return;
            }

            message ??= string.Empty;
            Assert.AreEqual(expected.IsHidden, actual.IsHidden, string.Format(message, "IsHidden"));
            Assert.AreEqual(expected.OutlineLevel, actual.OutlineLevel, string.Format(message, "OutlineLevel"));
            Assert.AreEqual(expected.Width, actual.Width, 1e-6, string.Format(message, "Width"));
        }

        public static void AreRowEquals(IXLRow expected, IXLRow actual, string message = null)
        {
            if (expected.Equals(actual))
            {
                return;
            }

            message ??= string.Empty;
            Assert.AreEqual(expected.IsHidden, actual.IsHidden, string.Format(message, "IsHidden"));
            Assert.AreEqual(expected.OutlineLevel, actual.OutlineLevel, string.Format(message, "OutlineLevel"));
            Assert.AreEqual(expected.Height, actual.Height, 1e-6, string.Format(message, "Height"));
        }

        public static void AreCellsStyleEquals(IXLStyle expected, IXLStyle actual, string message = null)
        {
            if (expected.Equals(actual))
            {
                return;
            }

            message ??= string.Empty;
            Assert.AreEqual(expected.Border, actual.Border, string.Format(message, "Border"));
            Assert.AreEqual(expected.Alignment, actual.Alignment, string.Format(message, "Alignment"));
            Assert.AreEqual(expected.DateFormat, actual.DateFormat, string.Format(message, "DateFormat"));
            Assert.AreEqual(expected.Fill, actual.Fill, string.Format(message, "Fill"));
            Assert.AreEqual(expected.Font, actual.Font, string.Format(message, "Font"));
            Assert.AreEqual(expected.NumberFormat, actual.NumberFormat, string.Format(message, "NumberFormat"));
            Assert.AreEqual(expected.Protection, actual.Protection, string.Format(message, "Protection"));
        }

        public static void AreCellsCommentEquals(IXLComment expected, IXLComment actual, string message = null)
        {
            if (expected == actual)
            {
                return;
            }

            message ??= string.Empty;
            Assert.AreEqual(expected.Text, actual.Text, string.Format(message, "Text"));
            Assert.AreEqual(expected.Author, actual.Author, string.Format(message, "Author"));
            Assert.AreEqual(expected.Count, actual.Count, string.Format(message, "Count"));
        }

        private static void AreMergedRangesEquals(IXLRanges expected, IXLRanges actual)
        {
            Assert.AreEqual(expected.Count(), actual.Count(), "Worksheet merged ranges count failed");
            IXLRange[] expectedArray = expected
                .OrderBy(r => r.RangeAddress.FirstAddress.RowNumber)
                .ThenBy(r => r.RangeAddress.FirstAddress.ColumnNumber)
                .ToArray();
            IXLRange[] actualArray = actual
                .OrderBy(r => r.RangeAddress.FirstAddress.RowNumber)
                .ThenBy(r => r.RangeAddress.FirstAddress.ColumnNumber)
                .ToArray();
            for (int i = 0; i < expectedArray.Length; i++)
            {
                Assert.AreEqual(expectedArray[i].RangeAddress.FirstAddress.RowNumber, actualArray[i].RangeAddress.FirstAddress.RowNumber, "Merge range first address row number failed");
                Assert.AreEqual(expectedArray[i].RangeAddress.FirstAddress.ColumnNumber, actualArray[i].RangeAddress.FirstAddress.ColumnNumber, "Merge range first address column number failed");
                Assert.AreEqual(expectedArray[i].RangeAddress.LastAddress.RowNumber, actualArray[i].RangeAddress.LastAddress.RowNumber, "Merge range last address row number failed");
                Assert.AreEqual(expectedArray[i].RangeAddress.LastAddress.ColumnNumber, actualArray[i].RangeAddress.LastAddress.ColumnNumber, "Merge range last address column number failed");
            }
        }

        public static void AreNamedRangesEquals(IXLNamedRanges expected, IXLNamedRanges actual)
        {
            Assert.AreEqual(expected.Count(), actual.Count(), "Worksheet named ranges count failed");
            foreach (IXLNamedRange expectedNamedRange in expected)
            {
                IXLNamedRange actualNamedRange = actual.NamedRange(expectedNamedRange.Name);
                Assert.AreEqual(expectedNamedRange.Comment, actualNamedRange.Comment, $"Named range {expectedNamedRange.Name} comment failed");
                Assert.AreEqual(expectedNamedRange.Ranges.Count, actualNamedRange.Ranges.Count, $"Named range {expectedNamedRange.Name} ranges count failed");
                for (int i = 0; i < expectedNamedRange.Ranges.Count; i++)
                {
                    IXLRange expectedRange = expectedNamedRange.Ranges.ElementAt(i);
                    IXLRange actualRange = actualNamedRange.Ranges.ElementAt(i);
                    Assert.AreEqual(expectedRange.FirstCell().Address, actualRange.FirstCell().Address, $"Named range {expectedNamedRange.Name} range {i + 1} first cell address failed");
                    Assert.AreEqual(expectedRange.LastCell().Address, actualRange.LastCell().Address, $"Named range {expectedNamedRange.Name} range {i + 1} last cell address failed");
                }
            }
        }

        public static void ArePageSetupEquals(IXLPageSetup expected, IXLPageSetup actual, string message = null)
        {
            if (expected == actual)
            {
                return;
            }

            message ??= string.Empty;

            Assert.AreEqual(expected.PagesTall, actual.PagesTall, string.Format(message, "PagesTall"));
            Assert.AreEqual(expected.PagesWide, actual.PagesWide, string.Format(message, "PagesWide"));
            Assert.AreEqual(expected.AlignHFWithMargins, actual.AlignHFWithMargins, string.Format(message, "AlignHFWithMargins"));
            Assert.AreEqual(expected.CenterVertically, actual.CenterVertically, string.Format(message, "CenterVertically"));
            Assert.AreEqual(expected.CenterHorizontally, actual.CenterHorizontally, string.Format(message, "CenterHorizontally"));
            Assert.AreEqual(expected.DifferentFirstPageOnHF, actual.DifferentFirstPageOnHF, string.Format(message, "DifferentFirstPageOnHF"));
            Assert.AreEqual(expected.DifferentOddEvenPagesOnHF, actual.DifferentOddEvenPagesOnHF, string.Format(message, "DifferentOddEvenPagesOnHF"));
            Assert.AreEqual(expected.DraftQuality, actual.DraftQuality, string.Format(message, "DraftQuality"));
            Assert.AreEqual(expected.FirstPageNumber, actual.FirstPageNumber, string.Format(message, "FirstPageNumber"));
            Assert.AreEqual(expected.VerticalDpi, actual.VerticalDpi, string.Format(message, "VerticalDpi"));
            Assert.AreEqual(expected.HorizontalDpi, actual.HorizontalDpi, string.Format(message, "HorizontalDpi"));
            Assert.AreEqual(expected.FirstRowToRepeatAtTop, actual.FirstRowToRepeatAtTop, string.Format(message, "FirstRowToRepeatAtTop"));
            Assert.AreEqual(expected.LastRowToRepeatAtTop, actual.LastRowToRepeatAtTop, string.Format(message, "LastRowToRepeatAtTop"));
            Assert.AreEqual(expected.FirstColumnToRepeatAtLeft, actual.FirstColumnToRepeatAtLeft, string.Format(message, "FirstColumnToRepeatAtLeft"));
            Assert.AreEqual(expected.LastColumnToRepeatAtLeft, actual.LastColumnToRepeatAtLeft, string.Format(message, "LastColumnToRepeatAtLeft"));
            Assert.AreEqual(expected.ScaleHFWithDocument, actual.ScaleHFWithDocument, string.Format(message, "ScaleHFWithDocument"));
            Assert.AreEqual(expected.ShowGridlines, actual.ShowGridlines, string.Format(message, "ShowGridlines"));
            Assert.AreEqual(expected.ShowRowAndColumnHeadings, actual.ShowRowAndColumnHeadings, string.Format(message, "ShowRowAndColumnHeadings"));
            Assert.AreEqual(expected.Scale, actual.Scale, string.Format(message, "Scale"));
            Assert.AreEqual(expected.PaperSize, actual.PaperSize, string.Format(message, "PaperSize"));
            Assert.AreEqual(expected.PageOrder, actual.PageOrder, string.Format(message, "PageOrder"));
            Assert.AreEqual(expected.PageOrientation, actual.PageOrientation, string.Format(message, "PageOrientation"));
            Assert.AreEqual(expected.BlackAndWhite, actual.BlackAndWhite, string.Format(message, "BlackAndWhite"));

            Assert.AreEqual(expected.ColumnBreaks.Count, actual.ColumnBreaks.Count, string.Format(message, "ColumnBreaks"));
            for (int i = 0; i < expected.ColumnBreaks.Count; i++)
            {
                Assert.AreEqual(expected.ColumnBreaks[i], actual.ColumnBreaks[i], string.Format(message, "ColumnBreaks"));
            }

            Assert.AreEqual(expected.RowBreaks.Count, actual.RowBreaks.Count, string.Format(message, "RowBreaks"));
            for (int i = 0; i < expected.RowBreaks.Count; i++)
            {
                Assert.AreEqual(expected.RowBreaks[i], actual.RowBreaks[i], string.Format(message, "RowBreaks"));
            }
        }
    }
}