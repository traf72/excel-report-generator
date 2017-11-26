using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ExcelReporter.Rendering.Panels.ExcelPanels;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourcePanelRenderTests
{
    [TestClass]
    public class DataSourcePanelDictionaryRenderTest
    {
        [TestMethod]
        public void TestRenderDictionary()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range1 = ws.Range(2, 2, 2, 3);
            range1.AddToNamed("TestRange", XLScope.Worksheet);

            IXLRange range2 = ws.Range(2, 5, 2, 6);
            range2.AddToNamed("TestRange2", XLScope.Worksheet);

            ws.Cell(1, 2).Value = "Key";
            ws.Cell(1, 3).Value = "Value";
            ws.Cell(2, 2).Value = "{di:Key}";
            ws.Cell(2, 3).Value = "{di:Value}";

            ws.Cell(1, 5).Value = "Key";
            ws.Cell(1, 6).Value = "Value";
            ws.Cell(2, 5).Value = "{di:Key}";
            ws.Cell(2, 6).Value = "{di:Value}";

            IDictionary<string, object> data1 = new DataSourcePanelDataProvider.TestDataProvider().GetDictionaryEnumerable().First();
            var panel1 = new ExcelDataSourcePanel(data1, ws.NamedRange("TestRange"), report);
            panel1.Render();

            IEnumerable<KeyValuePair<string, object>> data2 = new DataSourcePanelDataProvider.TestDataProvider().GetDictionaryEnumerable().First()
                .Select(x => new KeyValuePair<string, object>(x.Key, x.Value));
            var panel2 = new ExcelDataSourcePanel(data2, ws.NamedRange("TestRange2"), report);
            panel2.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelDictionaryRenderTest),
                nameof(TestRenderDictionary)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderDictionaryEnumerable()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range1 = ws.Range(2, 2, 2, 4);
            range1.AddToNamed("TestRange1", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Value}";
            ws.Cell(2, 4).Value = "{di:IsVip}";

            var panel1 = new ExcelDataSourcePanel("m:TestDataProvider:GetDictionaryEnumerable()", ws.NamedRange("TestRange1"), report);
            panel1.Render();

            var dictWihtDecimalValues = new List<IDictionary<string, decimal>>
            {
                new Dictionary<string, decimal> { ["Value"] = 25.7m },
                new Dictionary<string, decimal> { ["Value"] = 250.7m },
                new Dictionary<string, decimal> { ["Value"] = 2500.7m },
            };

            IXLRange range2 = ws.Range(2, 6, 2, 6);
            range2.AddToNamed("TestRange2", XLScope.Worksheet);

            ws.Cell(2, 6).Value = "{di:Value}";

            var panel2 = new ExcelDataSourcePanel(dictWihtDecimalValues, ws.NamedRange("TestRange2"), report);
            panel2.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelDictionaryRenderTest),
                nameof(TestRenderDictionaryEnumerable)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}