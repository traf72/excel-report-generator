using System.Linq;
using ClosedXML.Excel;
using ExcelReporter.Enums;
using ExcelReporter.Rendering.Panels.ExcelPanels;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourceDynamicPanelRenderTests
{
    [TestClass]
    public class DataSourceDynamicPanelDataSetRenderTest
    {
        [TestMethod]
        public void TestRenderDataSetWithEvents()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 4, 2);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{Headers}";
            ws.Cell(3, 2).Value = "{Data}";
            ws.Cell(4, 2).Value = "{Totals}";

            var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetAllCustomersDataSet()", ws.NamedRange("TestRange"), report)
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
            };
            panel.Render();

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

            var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetAllCustomersDataSet()", ws.NamedRange("TestRange"), report)
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
            };
            panel.Render();

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

            var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetAllCustomersDataSet()", ws.NamedRange("TestRange"), report)
            {
                BeforeRenderMethodName = "TestExcelDynamicPaneBeforeRender",
            };
            panel.Render();

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

            var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetAllCustomersDataSet()", ws.NamedRange("TestRange"), report)
            {
                AfterRenderMethodName = "TestExcelDynamicPaneAfterRender",
            };
            panel.Render();

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

            var panel = new ExcelDataSourceDynamicPanel("m:DataProvider:GetEmptyDataSet()", ws.NamedRange("TestRange"), report);
            panel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourceDynamicPanelDataSetRenderTest),
                nameof(TestRenderEmptyDataSet)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}