using ClosedXML.Excel;
using ExcelReporter.Implementations.Panels.Excel;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Implementations.Panels.Excel.PanelRenderTests.DataSourcePanelRenderTests
{
    [TestClass]
    public class DataSourcePanelDictionaryRenderTest
    {
        [TestMethod]
        public void TestRenderDictionary()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 2, 4);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Value}";
            ws.Cell(2, 4).Value = "{di:IsVip}";

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetDictionaryEnumerable()", ws.NamedRange("TestRange"), report);
            panel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelDictionaryRenderTest),
                nameof(TestRenderDictionary)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}