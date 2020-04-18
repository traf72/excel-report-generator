using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;
using NUnit.Framework;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourcePanelRenderTests
{
    
    public class DataSourcePanelDictionaryRenderTest
    {
        [Test]
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

            IDictionary<string, object> data1 = new DataProvider().GetDictionaryEnumerable().First();
            var panel1 = new ExcelDataSourcePanel(data1, ws.NamedRange("TestRange"), report, report.TemplateProcessor);
            panel1.Render();

            IEnumerable<KeyValuePair<string, object>> data2 = new DataProvider().GetDictionaryEnumerable().First()
                .Select(x => new KeyValuePair<string, object>(x.Key, x.Value));
            var panel2 = new ExcelDataSourcePanel(data2, ws.NamedRange("TestRange2"), report, report.TemplateProcessor);
            panel2.Render();

            Assert.AreEqual(ws.Range(2, 2, 4, 3), panel1.ResultRange);
            Assert.AreEqual(ws.Range(2, 5, 4, 6), panel2.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelDictionaryRenderTest),
                nameof(TestRenderDictionary)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [Test]
        public void TestRenderDictionaryEnumerable()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range1 = ws.Range(2, 2, 2, 4);
            range1.AddToNamed("TestRange1", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Value}";
            ws.Cell(2, 4).Value = "{di:IsVip}";

            var panel1 = new ExcelDataSourcePanel("m:DataProvider:GetDictionaryEnumerable()", ws.NamedRange("TestRange1"), report, report.TemplateProcessor);
            panel1.Render();

            var dictWithDecimalValues = new List<IDictionary<string, decimal>>
            {
                new Dictionary<string, decimal> { ["Value"] = 25.7m },
                new Dictionary<string, decimal> { ["Value"] = 250.7m },
                new Dictionary<string, decimal> { ["Value"] = 2500.7m },
            };

            IXLRange range2 = ws.Range(2, 6, 2, 6);
            range2.AddToNamed("TestRange2", XLScope.Worksheet);

            ws.Cell(2, 6).Value = "{di:Value}";

            var panel2 = new ExcelDataSourcePanel(dictWithDecimalValues, ws.NamedRange("TestRange2"), report, report.TemplateProcessor);
            panel2.Render();

            Assert.AreEqual(ws.Range(2, 2, 4, 4), panel1.ResultRange);
            Assert.AreEqual(ws.Range(2, 6, 4, 6), panel2.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelDictionaryRenderTest),
                nameof(TestRenderDictionaryEnumerable)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}