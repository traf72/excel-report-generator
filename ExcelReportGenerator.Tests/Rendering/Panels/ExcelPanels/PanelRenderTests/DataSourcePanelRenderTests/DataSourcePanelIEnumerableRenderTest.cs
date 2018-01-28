using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ClosedXML.Excel;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourcePanelRenderTests
{
    [TestClass]
    public class DataSourcePanelIEnumerableRenderTest
    {
        [TestMethod]
        public void TestRenderIEnumerableVerticalCellsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

            ws.Cell(2, 4).DataType = XLCellValues.Number;

            ws.Cell(2, 2).Value = "{m:Concat(di:Name, m:Format(di:Date, dd.MM.yyyy))}";
            ws.Cell(2, 3).Value = "{di:Date}";
            ws.Cell(2, 4).Value = "{m:Multiply(di:Sum, 5)}";
            ws.Cell(2, 5).Value = "{di:Contacts}";
            ws.Cell(3, 2).Value = "{Di:Contacts.Phone}";
            ws.Cell(3, 3).Value = "{DI:Contacts.Fax}";
            ws.Cell(3, 4).Value = "{p:StrParam}";

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(4, 1).Value = "{di:Name}";
            ws.Cell(1, 6).Value = "{di:Name}";
            ws.Cell(4, 6).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 6).Value = "{di:Name}";
            ws.Cell(1, 4).Value = "{di:Name}";
            ws.Cell(4, 4).Value = "{di:Name}";

            var panel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("TestRange"), report, report.TemplateProcessor);
            IXLRange resultRange = panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 7, 5), resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelIEnumerableRenderTest),
                nameof(TestRenderIEnumerableVerticalCellsShift)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderIEnumerableVerticalRowsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

            ws.Cell(2, 4).DataType = XLCellValues.Number;

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";
            ws.Cell(2, 4).Value = "{di:Sum}";
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

            var panel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("TestRange"), report, report.TemplateProcessor)
            {
                ShiftType = ShiftType.Row,
            };
            IXLRange resultRange = panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 7, 5), resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelIEnumerableRenderTest),
                nameof(TestRenderIEnumerableVerticalRowsShift)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderIEnumerableVerticalNoShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

            ws.Cell(3, 5).Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);

            ws.Cell(2, 4).DataType = XLCellValues.Number;

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";
            ws.Cell(2, 4).Value = "{di:Sum}";
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
            ws.Cell(8, 5).Value = "{di:Date}";

            ws.Cell(8, 5).Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            ws.Cell(8, 5).Style.Border.SetRightBorder(XLBorderStyleValues.Thin);
            ws.Cell(8, 5).Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
            ws.Cell(8, 5).Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

            ws.Cell(8, 5).Style.Border.SetTopBorderColor(XLColor.Red);
            ws.Cell(8, 5).Style.Border.SetRightBorderColor(XLColor.Red);
            ws.Cell(8, 5).Style.Border.SetBottomBorderColor(XLColor.Red);
            ws.Cell(8, 5).Style.Border.SetLeftBorderColor(XLColor.Red);

            var panel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("TestRange"), report, report.TemplateProcessor)
            {
                ShiftType = ShiftType.NoShift,
            };
            IXLRange resultRange = panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 7, 5), resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelIEnumerableRenderTest),
                nameof(TestRenderIEnumerableVerticalNoShift)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderIEnumerableHorizontalCellsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

            ws.Cell(2, 4).DataType = XLCellValues.Number;

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";
            ws.Cell(2, 4).Value = "{di:Sum}";
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

            var panel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("TestRange"), report, report.TemplateProcessor)
            {
                Type = PanelType.Horizontal,
            };
            IXLRange resultRange = panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 3, 13), resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelIEnumerableRenderTest),
                nameof(TestRenderIEnumerableHorizontalCellsShift)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderIEnumerableHorizontalRowsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

            ws.Cell(2, 4).DataType = XLCellValues.Number;

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";
            ws.Cell(2, 4).Value = "{di:Sum}";
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

            var panel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("TestRange"), report, report.TemplateProcessor)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.Row,
            };
            IXLRange resultRange = panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 3, 13), resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelIEnumerableRenderTest),
                nameof(TestRenderIEnumerableHorizontalRowsShift)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderIEnumerableHorizontalNoShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 3, 5);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
            range.FirstCell().Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

            ws.Cell(2, 4).DataType = XLCellValues.Number;

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";
            ws.Cell(2, 4).Value = "{di:Sum}";
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
            ws.Cell(2, 14).Value = "{di:Date}";

            var panel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("TestRange"), report, report.TemplateProcessor)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.NoShift,
            };
            IXLRange resultRange = panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 3, 13), resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelIEnumerableRenderTest),
                nameof(TestRenderIEnumerableHorizontalNoShift)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderIEnumerableOfInt()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 2, 2);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:di}";

            var panel = new ExcelDataSourcePanel(new[] {1, 2, 3, 4}, ws.NamedRange("TestRange"), report, report.TemplateProcessor);
            IXLRange resultRange = panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 5, 2), resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelIEnumerableRenderTest),
                nameof(TestRenderIEnumerableOfInt)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderIEnumerableOfString()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 2, 2);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:di}";

            var panel = new ExcelDataSourcePanel(new[] { "One", "Two", "Three", "Four" }, ws.NamedRange("TestRange"), report, report.TemplateProcessor);
            IXLRange resultRange = panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 5, 2), resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelIEnumerableRenderTest),
                nameof(TestRenderIEnumerableOfString)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestCancelPanelRender()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 2, 2);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:di}";

            var panel = new ExcelDataSourcePanel(new[] { 1, 2, 3, 4 }, ws.NamedRange("TestRange"), report, report.TemplateProcessor)
            {
                BeforeRenderMethodName = "CancelPanelRender",
            };
            IXLRange resultRange = panel.Render();

            Assert.AreEqual(range, resultRange);

            Assert.AreEqual(1, ws.CellsUsed().Count());
            Assert.AreEqual("{di:di}", ws.Cell(2, 2).Value);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestPanelRenderEvents()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 2, 3);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";

            var panel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("TestRange"), report, report.TemplateProcessor)
            {
                BeforeRenderMethodName = "TestExcelDataSourcePanelBeforeRender",
                AfterRenderMethodName = "TestExcelDataSourcePanelAfterRender",
                BeforeDataItemRenderMethodName = "TestExcelDataItemPanelBeforeRender",
                AfterDataItemRenderMethodName = "TestExcelDataItemPanelAfterRender",
            };
            IXLRange resultRange = panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 4, 2), resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelIEnumerableRenderTest),
                nameof(TestPanelRenderEvents)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        // Тестирование скорости рендеринга
        //[TestMethod]
        public void TestPanelRenderSpeed()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 2, 6);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";
            ws.Cell(2, 4).Value = "{di:Sum}";
            ws.Cell(2, 5).Value = "{di:Contacts.Phone}";
            ws.Cell(2, 6).Value = "{di:Contacts.Fax}";

            const int dataCount = 6000;
            IList<TestItem> data = new List<TestItem>(dataCount);
            for (int i = 0; i < dataCount; i++)
            {
                data.Add(new TestItem($"Name_{i}", DateTime.Now.AddHours(1), i + 10, new Contacts($"Phone_{i}", $"Fax_{i}")));
            }

            var panel = new ExcelDataSourcePanel(data, ws.NamedRange("TestRange"), report, report.TemplateProcessor)
            {
                //ShiftType = ShiftType.Row,
                //ShiftType = ShiftType.NoShift,
            };

            Stopwatch sw = Stopwatch.StartNew();

            panel.Render();

            sw.Stop();

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}