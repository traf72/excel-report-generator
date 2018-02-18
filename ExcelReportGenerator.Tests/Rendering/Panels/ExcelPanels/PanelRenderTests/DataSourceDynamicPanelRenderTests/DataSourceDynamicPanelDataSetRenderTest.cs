using System.Linq;
using ClosedXML.Excel;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourceDynamicPanelRenderTests
{
    [TestClass]
    public class DataSourceDynamicPanelDataSetRenderTest
    {
        [TestMethod]
        public void TestRenderDataSetWithEvents()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 5, 2);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{Headers}";
            ws.Cell(3, 2).Value = "{Numbers}";
            ws.Cell(4, 2).Value = "{Data}";
            ws.Cell(5, 2).Value = "{Totals}";

            var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetAllCustomersDataSet()", ws.NamedRange("TestRange"), report, report.TemplateProcessor)
            {
                BeforeHeadersRenderMethodName = "TestExcelDynamicPanelBeforeHeadersRender",
                AfterHeadersRenderMethodName = "TestExcelDynamicPanelAfterHeadersRender",
                BeforeNumbersRenderMethodName = "TestExcelDynamicPanelBeforeNumbersRender",
                AfterNumbersRenderMethodName = "TestExcelDynamicPanelAfterNumbersRender",
                BeforeDataTemplatesRenderMethodName = "TestExcelDynamicPanelBeforeDataTemplatesRender",
                AfterDataTemplatesRenderMethodName = "TestExcelDynamicPanelAfterDataTemplatesRender",
                BeforeDataRenderMethodName = "TestExcelDynamicPanelBeforeDataRender",
                AfterDataRenderMethodName = "TestExcelDynamicPanelAfterDataRender",
                BeforeDataItemRenderMethodName = "TestExcelDynamicPanelBeforeDataItemRender",
                AfterDataItemRenderMethodName = "TestExcelDynamicPanelAfterDataItemRender",
                BeforeTotalsTemplatesRenderMethodName = "TestExcelDynamicPanelBeforeTotalsTemplatesRender",
                AfterTotalsTemplatesRenderMethodName = "TestExcelDynamicPanelAfterTotalsTemplatesRender",
                BeforeTotalsRenderMethodName = "TestExcelDynamicPanelBeforeTotalsRender",
                AfterTotalsRenderMethodName = "TestExcelDynamicPaneAfterTotalsRender",
                GroupBy = "4",
            };
            IXLRange resultRange = panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 7, 8), resultRange);

            // Bug of ClosedXml - invalid determine of FirstCellUsed and LastCellUsed if merged ranges exist
            ws.Cell(1, 1).Value = "Stub";
            ws.Range(1, 1, 1, 1).Merge();
            ws.Cell(8, 9).Value = "Stub";
            ws.Range(8, 9, 8, 9).Merge();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourceDynamicPanelDataSetRenderTest),
                nameof(TestRenderDataSetWithEvents)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderDataSetWithEvents_HorizontalPanel()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 2, 4);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{Headers}";
            ws.Cell(2, 3).Value = "{Data}";
            ws.Cell(2, 4).Value = "{Totals}";

            var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetAllCustomersDataSet()", ws.NamedRange("TestRange"), report, report.TemplateProcessor)
            {
                BeforeHeadersRenderMethodName = "TestExcelDynamicPanelBeforeHeadersRender",
                AfterHeadersRenderMethodName = "TestExcelDynamicPanelAfterHeadersRender",
                BeforeDataTemplatesRenderMethodName = "TestExcelDynamicPanelBeforeDataTemplatesRender",
                AfterDataTemplatesRenderMethodName = "TestExcelDynamicPanelAfterDataTemplatesRender",
                BeforeDataRenderMethodName = "TestExcelDynamicPanelBeforeDataRender",
                AfterDataRenderMethodName = "TestExcelDynamicPanelAfterDataRender",
                BeforeDataItemRenderMethodName = "TestExcelDynamicPanelBeforeDataItemRender",
                AfterDataItemRenderMethodName = "TestExcelDynamicPanelAfterDataItemRender",
                BeforeTotalsTemplatesRenderMethodName = "TestExcelDynamicPanelBeforeTotalsTemplatesRender",
                AfterTotalsTemplatesRenderMethodName = "TestExcelDynamicPanelAfterTotalsTemplatesRender",
                BeforeTotalsRenderMethodName = "TestExcelDynamicPanelBeforeTotalsRender",
                AfterTotalsRenderMethodName = "TestExcelDynamicPaneAfterTotalsRender",
                Type = PanelType.Horizontal,
                GroupBy = "4",
            };
            IXLRange resultRange = panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 8, 6), resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourceDynamicPanelDataSetRenderTest),
                nameof(TestRenderDataSetWithEvents_HorizontalPanel)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestDynamicPanelBeforeRenderEvent()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 4, 2);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{Headers}";
            ws.Cell(3, 2).Value = "{Data}";
            ws.Cell(4, 2).Value = "{Totals}";

            var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetAllCustomersDataSet()", ws.NamedRange("TestRange"), report, report.TemplateProcessor)
            {
                BeforeRenderMethodName = "TestExcelDynamicPaneBeforeRender",
            };
            IXLRange resultRange = panel.Render();

            Assert.AreEqual(range, resultRange);

            Assert.AreEqual(3, ws.CellsUsed().Count());
            Assert.AreEqual("CanceledHeaders", ws.Cell(2, 2).Value);
            Assert.AreEqual("CanceledData", ws.Cell(3, 2).Value);
            Assert.AreEqual("CanceledTotals", ws.Cell(4, 2).Value);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestDynamicPanelAfterRenderEvent()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 4, 2);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{Headers}";
            ws.Cell(3, 2).Value = "{Data}";
            ws.Cell(4, 2).Value = "{Totals}";

            var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetAllCustomersDataSet()", ws.NamedRange("TestRange"), report, report.TemplateProcessor)
            {
                AfterRenderMethodName = "TestExcelDynamicPaneAfterRender",
            };
            IXLRange resultRange = panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 6, 7), resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourceDynamicPanelDataSetRenderTest),
                nameof(TestDynamicPanelAfterRenderEvent)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderEmptyDataSet()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 4, 2);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{Headers}";
            ws.Cell(3, 2).Value = "{Data}";
            ws.Cell(4, 2).Value = "{Totals}";

            var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetEmptyDataSet()", ws.NamedRange("TestRange"), report, report.TemplateProcessor);
            IXLRange resultRange = panel.Render();

            Assert.AreEqual(ws.Range(2, 2, 3, 7), resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourceDynamicPanelDataSetRenderTest),
                nameof(TestRenderEmptyDataSet)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}