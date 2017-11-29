using ClosedXML.Excel;
using ExcelReporter.Enums;
using ExcelReporter.Rendering.Panels.ExcelPanels;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourcePanelRenderTests
{
    [TestClass]
    public class DataSourcePanelRender_WithGrouping_MultipleChildrenInOneParent
    {
        [TestMethod]
        public void Test_TwoChildren_Vertical()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 6, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange child1 = ws.Range(4, 2, 4, 5);
            child1.AddToNamed("ChildRange1", XLScope.Worksheet);

            IXLRange child2 = ws.Range(6, 2, 6, 5);
            child2.AddToNamed("ChildRange2", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";

            ws.Cell(3, 3).Value = "Field1";
            ws.Cell(3, 3).Style.Font.Bold = true;
            ws.Cell(3, 4).Value = "Field2";
            ws.Cell(3, 4).Style.Font.Bold = true;
            ws.Cell(4, 3).Value = "{di:Field1}";
            ws.Cell(4, 4).Value = "{di:Field2}";
            ws.Cell(5, 5).Value = "Number";
            ws.Cell(5, 5).Style.Font.Bold = true;
            ws.Cell(6, 5).Value = "{di:di}";

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report);
            var childPanel1 = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange1"), report)
            {
                Parent = parentPanel,
            };
            var childPanel2 = new ExcelDataSourcePanel("di:ChildrenPrimitive", ws.NamedRange("ChildRange2"), report)
            {
                Parent = parentPanel,
            };
            parentPanel.Children = new[] { childPanel1, childPanel2 };
            parentPanel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_MultipleChildrenInOneParent),
                nameof(Test_TwoChildren_Vertical)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_TwoChildren_Horizontal()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 5, 6);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange child1 = ws.Range(2, 4, 5, 4);
            child1.AddToNamed("ChildRange1", XLScope.Worksheet);

            IXLRange child2 = ws.Range(2, 6, 5, 6);
            child2.AddToNamed("ChildRange2", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";

            ws.Cell(3, 3).Value = "Field1";
            ws.Cell(3, 3).Style.Font.Bold = true;
            ws.Cell(4, 3).Value = "Field2";
            ws.Cell(4, 3).Style.Font.Bold = true;
            ws.Cell(3, 4).Value = "{di:Field1}";
            ws.Cell(4, 4).Value = "{di:Field2}";
            ws.Cell(5, 5).Value = "Number";
            ws.Cell(5, 5).Style.Font.Bold = true;
            ws.Cell(5, 6).Value = "{di:di}";

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report);
            var childPanel1 = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange1"), report)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
            };
            var childPanel2 = new ExcelDataSourcePanel("di:ChildrenPrimitive", ws.NamedRange("ChildRange2"), report)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
            };
            parentPanel.Children = new[] { childPanel1, childPanel2 };
            parentPanel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_MultipleChildrenInOneParent),
                nameof(Test_TwoChildren_Horizontal)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}