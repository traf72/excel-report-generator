using System.Linq;
using ClosedXML.Excel;
using ExcelReportGenerator.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReportGenerator.Tests.Extensions
{
    [TestClass]
    public class XLRangeBaseExtensionsTest
    {
        [TestMethod]
        public void TestCellsUsedWithoutFormulas()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");
            ws.Cell(70, 100).Value = "Value";
            ws.Cell(10, 15).Value = "Value2";
            ws.Cell(10, 15).Active = true;
            ws.Cell(10, 20).FormulaA1 = "=ROW()";
            ws.Cell(10, 30).FormulaR1C1 = "=COLUMN()";
            ws.Cell(20, 30).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Cell(25, 30).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Cell(25, 30).FormulaA1 = "=A1+B2";

            Assert.AreEqual(5, ws.CellsUsed().Count());
            Assert.AreEqual(6, ws.CellsUsed(true).Count());
            Assert.AreEqual(2, ws.CellsUsedWithoutFormulas().Count());
            Assert.AreEqual(3, ws.CellsUsedWithoutFormulas(true).Count());
            Assert.AreEqual(2, ws.CellsUsed(c => c.Active || c.Style.Border.TopBorder == XLBorderStyleValues.Thin).Count());
            Assert.AreEqual(1, ws.CellsUsedWithoutFormulas(c => c.Active || c.Style.Border.TopBorder == XLBorderStyleValues.Thin).Count());
            Assert.AreEqual(3, ws.CellsUsed(true, c => c.Active || c.Style.Border.TopBorder == XLBorderStyleValues.Thin).Count());
            Assert.AreEqual(2, ws.CellsUsedWithoutFormulas(true, c => c.Active || c.Style.Border.TopBorder == XLBorderStyleValues.Thin).Count());
        }

        [TestMethod]
        public void TestCellsWithoutFormulas()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLRange range = ws.Range(1, 1, 10, 10);

            range.Cell(1, 1).Value = "Value";
            range.Cell(2, 2).Value = "Value2";
            range.Cell(2, 2).Active = true;
            range.Cell(3, 3).FormulaA1 = "=ROW()";
            range.Cell(4, 4).FormulaR1C1 = "=COLUMN()";
            range.Cell(5, 5).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            range.Cell(6, 6).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            range.Cell(6, 6).FormulaA1 = "=A1+B2";

            Assert.AreEqual(100, range.Cells().Count());
            Assert.AreEqual(97, range.CellsWithoutFormulas().Count());
            Assert.AreEqual(3, range.Cells(c => c.Active || c.Style.Border.TopBorder == XLBorderStyleValues.Thin).Count());
            Assert.AreEqual(2, range.CellsWithoutFormulas(c => c.Active || c.Style.Border.TopBorder == XLBorderStyleValues.Thin).Count());
        }
    }
}