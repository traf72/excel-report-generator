using ClosedXML.Excel;
using ExcelReporter.Enums;
using ExcelReporter.Rendering.Panels.ExcelPanels;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourcePanelRenderTests
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

            var panel = new ExcelDataSourcePanel("m:DataProvider:GetSingleItem()", ws.NamedRange("TestRange"), report);
            panel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelOneItemRenderTest),
                nameof(TestRenderOneItemVerticalCellsShift)), ws.Workbook);

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

            var panel = new ExcelDataSourcePanel("m:DataProvider:GetSingleItem()", ws.NamedRange("TestRange"), report)
            {
                ShiftType = ShiftType.Row,
            };
            panel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelOneItemRenderTest),
                nameof(TestRenderOneItemVerticalRowShift)), ws.Workbook);

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

            var panel = new ExcelDataSourcePanel("m:DataProvider:GetSingleItem()", ws.NamedRange("TestRange"), report)
            {
                ShiftType = ShiftType.NoShift,
            };
            panel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelOneItemRenderTest),
                nameof(TestRenderOneItemVerticalNoShift)), ws.Workbook);

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

            var panel = new ExcelDataSourcePanel("m:DataProvider:GetSingleItem()", ws.NamedRange("TestRange"), report)
            {
                Type = PanelType.Horizontal,
            };
            panel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelOneItemRenderTest),
                nameof(TestRenderOneItemHorizontalCellsShift)), ws.Workbook);

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

            var panel = new ExcelDataSourcePanel("m:DataProvider:GetSingleItem()", ws.NamedRange("TestRange"), report)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.Row,
            };
            panel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelOneItemRenderTest),
                nameof(TestRenderOneItemHorizontalRowShift)), ws.Workbook);

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

            var panel = new ExcelDataSourcePanel("m:DataProvider:GetSingleItem()", ws.NamedRange("TestRange"), report)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.NoShift,
            };
            panel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelOneItemRenderTest),
                nameof(TestRenderOneItemHorizontalNoShift)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}