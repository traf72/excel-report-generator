using ClosedXML.Excel;
using ExcelReporter.Rendering.Panels.ExcelPanels;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReporter.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourceDynamicPanelRenderTests
{
    [TestClass]
    public class DataSourceDynamicPanelDictionaryRenderTest
    {
        [TestMethod]
        public void TestRenderDictionary()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range1 = ws.Range(2, 2, 4, 2);
            range1.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{Headers}";
            ws.Cell(3, 2).Value = "{Data}";
            ws.Cell(4, 2).Value = "{Totals}";

            IXLRange range2 = ws.Range(7, 2, 9, 2);
            range2.AddToNamed("TestRange2", XLScope.Worksheet);

            ws.Cell(7, 2).Value = "{Headers}";
            ws.Cell(8, 2).Value = "{Data}";
            ws.Cell(9, 2).Value = "{Totals}";

            IDictionary<string, object> data1 = new DataProvider().GetDictionaryEnumerable().First();
            var panel1 = new ExcelDataSourceDynamicPanel(data1, ws.NamedRange("TestRange"), report);
            panel1.Render();

            IEnumerable<KeyValuePair<string, object>> data2 = new DataProvider().GetDictionaryEnumerable().First()
                .Select(x => new KeyValuePair<string, object>(x.Key, x.Value));
            var panel2 = new ExcelDataSourceDynamicPanel(data2, ws.NamedRange("TestRange2"), report);
            panel2.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourceDynamicPanelDictionaryRenderTest),
                nameof(TestRenderDictionary)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderEmptyDictionary()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range1 = ws.Range(2, 2, 4, 2);
            range1.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{Headers}";
            ws.Cell(3, 2).Value = "{Data}";
            ws.Cell(4, 2).Value = "{Totals}";

            IXLRange range2 = ws.Range(7, 2, 9, 2);
            range2.AddToNamed("TestRange2", XLScope.Worksheet);

            ws.Cell(7, 2).Value = "{Headers}";
            ws.Cell(8, 2).Value = "{Data}";
            ws.Cell(9, 2).Value = "{Totals}";

            IDictionary<string, object> data1 = new Dictionary<string, object>();
            var panel1 = new ExcelDataSourceDynamicPanel(data1, ws.NamedRange("TestRange"), report);
            panel1.Render();

            IEnumerable<KeyValuePair<string, object>> data2 = new List<KeyValuePair<string, object>>();
            var panel2 = new ExcelDataSourceDynamicPanel(data2, ws.NamedRange("TestRange2"), report);
            panel2.Render();

            Assert.AreEqual(4, ws.CellsUsed().Count());
            Assert.AreEqual(ws.Cell(2, 2).Value, "Key");
            Assert.AreEqual(ws.Cell(2, 3).Value, "Value");
            Assert.AreEqual(ws.Cell(6, 2).Value, "Key");
            Assert.AreEqual(ws.Cell(6, 3).Value, "Value");

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderDictionaryEnumerable()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range1 = ws.Range(2, 2, 4, 2);
            range1.AddToNamed("TestRange1", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{Headers}";
            ws.Cell(3, 2).Value = "{Data}";
            ws.Cell(4, 2).Value = "{Totals}";

            var panel1 = new ExcelDataSourceDynamicPanel("m:DataProvider:GetDictionaryEnumerable()", ws.NamedRange("TestRange1"), report);
            panel1.Render();

            var dictWihtDecimalValues = new List<IDictionary<string, decimal>>
            {
                new Dictionary<string, decimal> { ["Value"] = 25.7m },
                new Dictionary<string, decimal> { ["Value"] = 250.7m },
                new Dictionary<string, decimal> { ["Value"] = 2500.7m },
            };

            IXLRange range2 = ws.Range(7, 2, 9, 2);
            range2.AddToNamed("TestRange2", XLScope.Worksheet);

            ws.Cell(7, 2).Value = "{Headers}";
            ws.Cell(8, 2).Value = "{Data}";
            ws.Cell(9, 2).Value = "{Totals}";

            var panel2 = new ExcelDataSourceDynamicPanel(dictWihtDecimalValues, ws.NamedRange("TestRange2"), report);
            panel2.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourceDynamicPanelDictionaryRenderTest),
                nameof(TestRenderDictionaryEnumerable)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderEmptyDictionaryEnumerable()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range1 = ws.Range(2, 2, 4, 2);
            range1.AddToNamed("TestRange1", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{Headers}";
            ws.Cell(3, 2).Value = "{Data}";
            ws.Cell(4, 2).Value = "{Totals}";

            var panel1 = new ExcelDataSourceDynamicPanel(new List<IDictionary<string, decimal>>(), ws.NamedRange("TestRange1"), report);
            panel1.Render();

            Assert.AreEqual(0, ws.CellsUsed().Count());

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}