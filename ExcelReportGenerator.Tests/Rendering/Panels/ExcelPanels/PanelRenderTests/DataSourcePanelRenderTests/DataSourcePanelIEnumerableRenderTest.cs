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

            ws.Cell(2, 4).DataType = XLDataType.Number;

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
            panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 7, 5), panel.ResultRange);

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

            ws.Cell(2, 4).DataType = XLDataType.Number;

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
            panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 7, 5), panel.ResultRange);

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

            ws.Cell(2, 4).DataType = XLDataType.Number;

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
            panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 7, 5), panel.ResultRange);

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

            ws.Cell(2, 4).DataType = XLDataType.Number;

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
            panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 3, 13), panel.ResultRange);

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

            ws.Cell(2, 4).DataType = XLDataType.Number;

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
            panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 3, 13), panel.ResultRange);

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

            ws.Cell(2, 4).DataType = XLDataType.Number;

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
            panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 3, 13), panel.ResultRange);

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
            panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 5, 2), panel.ResultRange);

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
            panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 5, 2), panel.ResultRange);

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
            panel.Render();

            Assert.AreEqual(range, panel.ResultRange);

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
            panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 4, 2), panel.ResultRange);

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

            //Stopwatch sw2 = Stopwatch.StartNew();

            //report.Workbook.SaveAs("test.xlsx");

            //sw2.Stop();
        }

        // Тестирование скорости рендеринга
        //[TestMethod]
        public void TestPanelRenderSpeedWithHierarchy()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 4, 6);
            range.AddToNamed("ParentRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";
            ws.Cell(2, 4).Value = "{di:Sum}";
            ws.Cell(2, 5).Value = "{di:Contacts.Phone}";
            ws.Cell(2, 6).Value = "{di:Contacts.Fax}";

            IXLRange child = ws.Range(3, 2, 3, 6);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            ws.Cell(3, 2).Value = "{di:Field1}";
            ws.Cell(3, 3).Value = "{di:Field2}";

            IXLRange total = ws.Range(4, 2, 4, 6);
            total.AddToNamed("TotalRange", XLScope.Worksheet);

            ws.Cell(4, 5).Value = "{Max(di:Field1)}";
            ws.Cell(4, 6).Value = "{Min(di:Field2)}";

            const int dataCount = 100;
            IList<TestItem> data = new List<TestItem>(dataCount);
            for (int i = 0; i < dataCount; i++)
            {
                data.Add(new TestItem($"Name_{i}", DateTime.Now.AddHours(1), i + 1, new Contacts($"Phone_{i}", $"Fax_{i}")));
            }

            var parentPanel = new ExcelDataSourcePanel(data, ws.NamedRange("ParentRange"), report, report.TemplateProcessor)
            {
                //ShiftType = ShiftType.Row,
                //ShiftType = ShiftType.NoShift,
            };

            //var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildrenProportionally(di:di)", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildrenRandom(10, 20)", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                //ShiftType = ShiftType.Row,
                //ShiftType = ShiftType.NoShift,
                Parent = parentPanel,
            };

            //var totalPanel = new ExcelTotalsPanel("m:DataProvider:GetChildrenProportionally(di:di)", ws.NamedRange("TotalRange"), report, report.TemplateProcessor)
            var totalPanel = new ExcelTotalsPanel("m:DataProvider:GetChildrenRandom(10, 20)", ws.NamedRange("TotalRange"), report, report.TemplateProcessor)
            {
                //ShiftType = ShiftType.Row,
                //ShiftType = ShiftType.NoShift,
                Parent = parentPanel,
            };

            parentPanel.Children.Add(childPanel);
            parentPanel.Children.Add(totalPanel);

            Stopwatch sw = Stopwatch.StartNew();

            parentPanel.Render();

            sw.Stop();

            //Stopwatch sw2 = Stopwatch.StartNew();

            //report.Workbook.SaveAs("test.xlsx");

            //sw2.Stop();
        }

        // Тестирование скорости рендеринга
        //[TestMethod]
        public void TestPanelRenderSpeedWithMultiHierarchy()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 4, 6);
            range.AddToNamed("ParentRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";
            ws.Cell(2, 4).Value = "{di:Sum}";
            ws.Cell(2, 5).Value = "{di:Contacts.Phone}";
            ws.Cell(2, 6).Value = "{di:Contacts.Fax}";

            IXLRange child1 = ws.Range(3, 2, 4, 6);
            child1.AddToNamed("ChildRange1", XLScope.Worksheet);

            ws.Cell(3, 2).Value = "{di:Field1}";
            ws.Cell(3, 3).Value = "{di:Field2}";

            IXLRange child2 = ws.Range(4, 2, 4, 6);
            child2.AddToNamed("ChildRange2", XLScope.Worksheet);

            ws.Cell(4, 5).Value = "{di:Field1}";
            ws.Cell(4, 6).Value = "{di:Field2}";

            const int dataCount = 50;
            IList<TestItem> data = new List<TestItem>(dataCount);
            for (int i = 0; i < dataCount; i++)
            {
                data.Add(new TestItem($"Name_{i}", DateTime.Now.AddHours(1), i + 1, new Contacts($"Phone_{i}", $"Fax_{i}")));
            }

            var parentPanel = new ExcelDataSourcePanel(data, ws.NamedRange("ParentRange"), report, report.TemplateProcessor)
            {
                //ShiftType = ShiftType.Row,
                //ShiftType = ShiftType.NoShift,
            };

            var childPanel1 = new ExcelDataSourcePanel("m:DataProvider:GetChildrenRandom(4, 6)", ws.NamedRange("ChildRange1"), report, report.TemplateProcessor)
            {
                //ShiftType = ShiftType.Row,
                //ShiftType = ShiftType.NoShift,
                Parent = parentPanel,
            };

            var childPanel2 = new ExcelDataSourcePanel("m:DataProvider:GetChildrenRandom(10, 15)", ws.NamedRange("ChildRange2"), report, report.TemplateProcessor)
            {
                //ShiftType = ShiftType.Row,
                //ShiftType = ShiftType.NoShift,
                Parent = childPanel1,
            };

            parentPanel.Children.Add(childPanel1);
            childPanel1.Children.Add(childPanel2);

            Stopwatch sw = Stopwatch.StartNew();

            parentPanel.Render();

            sw.Stop();

            //Stopwatch sw2 = Stopwatch.StartNew();

            //report.Workbook.SaveAs("test.xlsx");

            //sw2.Stop();
        }
    }
}