using ClosedXML.Excel;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;
using NUnit.Framework;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourcePanelRenderTests
{
    
    public class DataSourcePanelSingleItemRenderTest
    {
        [Test]
        public void TestRenderSingleItemVerticalCellsShift()
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

            var panel = new ExcelDataSourcePanel("m:DataProvider:GetSingleItem()", ws.NamedRange("TestRange"), report, report.TemplateProcessor);
            panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 3, 5), panel.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelSingleItemRenderTest),
                nameof(TestRenderSingleItemVerticalCellsShift)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [Test]
        public void TestRenderSingleItemVerticalRowShift()
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

            var panel = new ExcelDataSourcePanel("m:DataProvider:GetSingleItem()", ws.NamedRange("TestRange"), report, report.TemplateProcessor)
            {
                ShiftType = ShiftType.Row,
            };
            panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 3, 5), panel.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelSingleItemRenderTest),
                nameof(TestRenderSingleItemVerticalRowShift)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [Test]
        public void TestRenderSingleItemVerticalNoShift()
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

            var panel = new ExcelDataSourcePanel("m:DataProvider:GetSingleItem()", ws.NamedRange("TestRange"), report, report.TemplateProcessor)
            {
                ShiftType = ShiftType.NoShift,
            };
            panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 3, 5), panel.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelSingleItemRenderTest),
                nameof(TestRenderSingleItemVerticalNoShift)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [Test]
        public void TestRenderSingleItemHorizontalCellsShift()
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

            var panel = new ExcelDataSourcePanel("m:DataProvider:GetSingleItem()", ws.NamedRange("TestRange"), report, report.TemplateProcessor)
            {
                Type = PanelType.Horizontal,
            };
            panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 3, 5), panel.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelSingleItemRenderTest),
                nameof(TestRenderSingleItemHorizontalCellsShift)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [Test]
        public void TestRenderSingleItemHorizontalRowShift()
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

            var panel = new ExcelDataSourcePanel("m:DataProvider:GetSingleItem()", ws.NamedRange("TestRange"), report, report.TemplateProcessor)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.Row,
            };
            panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 3, 5), panel.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelSingleItemRenderTest),
                nameof(TestRenderSingleItemHorizontalRowShift)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [Test]
        public void TestRenderSingleItemHorizontalNoShift()
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

            var panel = new ExcelDataSourcePanel("m:DataProvider:GetSingleItem()", ws.NamedRange("TestRange"), report, report.TemplateProcessor)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.NoShift,
            };
            panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 3, 5), panel.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelSingleItemRenderTest),
                nameof(TestRenderSingleItemHorizontalNoShift)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}