using System.Linq;
using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.CustomAsserts
{
    public class ExcelAssert
    {
        public static void AreWorksheetsContentEquals(IXLWorksheet expected, IXLWorksheet actual)
        {
            if (expected == actual)
            {
                return;
            }

            Assert.AreEqual(expected.CellsUsed(true).Count(), actual.CellsUsed(true).Count(), "Cells used count failed");

            IXLCell expectedFirstCellUsed = expected.FirstCellUsed(true);
            IXLCell actualFirstCellUsed = actual.FirstCellUsed(true);
            Assert.AreEqual(expectedFirstCellUsed.Address, actualFirstCellUsed.Address, "First cell used failed");
            IXLCell expectedLastCellUsed = expected.LastCellUsed(true);
            IXLCell actualLastCellUsed = actual.LastCellUsed(true);
            Assert.AreEqual(expectedLastCellUsed.Address, actualLastCellUsed.Address, "Last cell used failed");

            IXLRange range = expected.Range(expectedFirstCellUsed, expectedLastCellUsed);
            foreach (IXLCell expectedCell in range.Cells())
            {
                IXLCell actualCell = actual.Cell(expectedCell.Address);
                Assert.AreEqual(expectedCell.Value, actualCell.Value, $"Cell {expectedCell.Address} Value failed.");
                Assert.AreEqual(expectedCell.DataType, actualCell.DataType, $"Cell {expectedCell.Address} DataType failed.");
                Assert.AreEqual(expectedCell.Active, actualCell.Active, $"Cell {expectedCell.Address} Active failed.");
                AreColumnsEquals(expectedCell.WorksheetColumn(), actualCell.WorksheetColumn(), $"Column {actualCell.WorksheetColumn().RangeAddress} {{0}} failed.");
                AreRowEquals(expectedCell.WorksheetRow(), actualCell.WorksheetRow(), $"Row {actualCell.WorksheetRow().RangeAddress} {{0}} failed.");
                AreCellsStyleEquals(expectedCell.Style, actualCell.Style, $"Cell {expectedCell.Address} Style {{0}} failed.");
                AreCellsCommentEquals(expectedCell.Comment, actualCell.Comment, $"Cell {expectedCell.Address} Comment {{0}} failed.");
            }

            Assert.AreEqual(expected.NamedRanges.Count(), actual.NamedRanges.Count(), "Worksheet named ranges count failed");
            AreNamedRangesEquals(expected.NamedRanges, actual.NamedRanges);
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

            message = message ?? string.Empty;
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

            message = message ?? string.Empty;
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

            message = message ?? string.Empty;
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

            message = message ?? string.Empty;
            Assert.AreEqual(expected.Text, actual.Text, string.Format(message, "Text"));
            Assert.AreEqual(expected.Author, actual.Author, string.Format(message, "Author"));
            Assert.AreEqual(expected.Count, actual.Count, string.Format(message, "Count"));
        }

        public static void AreNamedRangesEquals(IXLNamedRanges expected, IXLNamedRanges actual)
        {
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
    }
}