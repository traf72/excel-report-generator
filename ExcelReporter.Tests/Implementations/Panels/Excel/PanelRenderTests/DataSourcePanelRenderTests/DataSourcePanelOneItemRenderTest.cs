using ClosedXML.Excel;
using ExcelReporter.Enums;
using ExcelReporter.Implementations.Panels.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace ExcelReporter.Tests.Implementations.Panels.Excel.PanelRenderTests.DataSourcePanelRenderTests
{
    [TestClass]
    public class DataSourcePanelOneItemRenderTest
    {
        [TestMethod]
        public void TestRenderOneItemVerticalCellsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";
            ws.Cell(2, 4).Value = "{m:Multiply(di:Sum, 5)}";
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

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetSingleItem()", ws.NamedRange("TestRange"), report);
            panel.Render();

            Assert.AreEqual(15, ws.CellsUsed().Count());
            Assert.AreEqual("Test", ws.Cell(2, 2).Value);
            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(2, 3).Value);
            Assert.AreEqual(278.8, ws.Cell(2, 4).Value);
            Assert.AreEqual("15_345", ws.Cell(2, 5).Value);
            Assert.AreEqual(15d, ws.Cell(3, 2).Value);
            Assert.AreEqual(345d, ws.Cell(3, 3).Value);
            Assert.AreEqual("String parameter", ws.Cell(3, 4).Value);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 4).Value);

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderOneItemVerticalRowShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";
            ws.Cell(2, 4).Value = "{m:Multiply(di:Sum, 5)}";
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

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetSingleItem()", ws.NamedRange("TestRange"), report)
            {
                ShiftType = ShiftType.Row,
            };
            panel.Render();

            Assert.AreEqual(15, ws.CellsUsed().Count());
            Assert.AreEqual("Test", ws.Cell(2, 2).Value);
            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(2, 3).Value);
            Assert.AreEqual(278.8, ws.Cell(2, 4).Value);
            Assert.AreEqual("15_345", ws.Cell(2, 5).Value);
            Assert.AreEqual(15d, ws.Cell(3, 2).Value);
            Assert.AreEqual(345d, ws.Cell(3, 3).Value);
            Assert.AreEqual("String parameter", ws.Cell(3, 4).Value);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 4).Value);

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderOneItemVerticalNoShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";
            ws.Cell(2, 4).Value = "{m:Multiply(di:Sum, 5)}";
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

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetSingleItem()", ws.NamedRange("TestRange"), report)
            {
                ShiftType = ShiftType.NoShift,
            };
            panel.Render();

            Assert.AreEqual(15, ws.CellsUsed().Count());
            Assert.AreEqual("Test", ws.Cell(2, 2).Value);
            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(2, 3).Value);
            Assert.AreEqual(278.8, ws.Cell(2, 4).Value);
            Assert.AreEqual("15_345", ws.Cell(2, 5).Value);
            Assert.AreEqual(15d, ws.Cell(3, 2).Value);
            Assert.AreEqual(345d, ws.Cell(3, 3).Value);
            Assert.AreEqual("String parameter", ws.Cell(3, 4).Value);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 4).Value);

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderOneItemHorizontalCellsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";
            ws.Cell(2, 4).Value = "{m:Multiply(di:Sum, 5)}";
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

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetSingleItem()", ws.NamedRange("TestRange"), report)
            {
                Type = PanelType.Horizontal,
            };
            panel.Render();

            Assert.AreEqual(15, ws.CellsUsed().Count());
            Assert.AreEqual("Test", ws.Cell(2, 2).Value);
            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(2, 3).Value);
            Assert.AreEqual(278.8, ws.Cell(2, 4).Value);
            Assert.AreEqual("15_345", ws.Cell(2, 5).Value);
            Assert.AreEqual(15d, ws.Cell(3, 2).Value);
            Assert.AreEqual(345d, ws.Cell(3, 3).Value);
            Assert.AreEqual("String parameter", ws.Cell(3, 4).Value);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 4).Value);

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderOneItemHorizontalRowShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";
            ws.Cell(2, 4).Value = "{m:Multiply(di:Sum, 5)}";
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

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetSingleItem()", ws.NamedRange("TestRange"), report)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.Row,
            };
            panel.Render();

            Assert.AreEqual(15, ws.CellsUsed().Count());
            Assert.AreEqual("Test", ws.Cell(2, 2).Value);
            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(2, 3).Value);
            Assert.AreEqual(278.8, ws.Cell(2, 4).Value);
            Assert.AreEqual("15_345", ws.Cell(2, 5).Value);
            Assert.AreEqual(15d, ws.Cell(3, 2).Value);
            Assert.AreEqual(345d, ws.Cell(3, 3).Value);
            Assert.AreEqual("String parameter", ws.Cell(3, 4).Value);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 4).Value);

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderOneItemHorizontalNoShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";
            ws.Cell(2, 4).Value = "{m:Multiply(di:Sum, 5)}";
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

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetSingleItem()", ws.NamedRange("TestRange"), report)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.NoShift,
            };
            panel.Render();

            Assert.AreEqual(15, ws.CellsUsed().Count());
            Assert.AreEqual("Test", ws.Cell(2, 2).Value);
            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(2, 3).Value);
            Assert.AreEqual(278.8, ws.Cell(2, 4).Value);
            Assert.AreEqual("15_345", ws.Cell(2, 5).Value);
            Assert.AreEqual(15d, ws.Cell(3, 2).Value);
            Assert.AreEqual(345d, ws.Cell(3, 3).Value);
            Assert.AreEqual("String parameter", ws.Cell(3, 4).Value);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 4).Value);

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}