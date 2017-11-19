using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Implementations.Panels.Excel.PanelRenderTests.DataSourcePanelRenderTests
{
    [TestClass]
    public class DataSourcePanelRender_WithGrouping_MultipleChildrenInOneParent
    {
        //[TestMethod]
        public void Test_TwoChildren_Vertical()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            //IXLRange parentRange = ws.Range(2, 2, 5, 3);
            //parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            //IXLRange child = ws.Range(2, 2, 5, 2);
            //child.AddToNamed("ChildRange", XLScope.Worksheet);

            //child.Range(2, 1, 4, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            //child.Range(2, 1, 4, 1).Style.Border.OutsideBorderColor = XLColor.Red;

            //parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            //parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            var usedCell = ws.LastCellUsed(true);

            //ws.Cell(2, 3).Value = "{di:Name}";
            //ws.Cell(3, 3).Value = "{di:Date}";

            //ws.Cell(3, 2).Value = "{di:Field1}";
            //ws.Cell(4, 2).Value = "{di:Field2}";
            //ws.Cell(5, 2).Value = "{di:parent:Sum}";

            //ws.Cell(1, 1).Value = "{di:Name}";
            //ws.Cell(1, 3).Value = "{di:Name}";
            //ws.Cell(1, 4).Value = "{di:Name}";
            //ws.Cell(3, 1).Value = "{di:Name}";
            //ws.Cell(3, 4).Value = "{di:Name}";
            //ws.Cell(6, 1).Value = "{di:Name}";
            //ws.Cell(6, 3).Value = "{di:Name}";
            //ws.Cell(6, 4).Value = "{di:Name}";

            //var parentPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            //{
            //    Type = PanelType.Horizontal,
            //};
            //var childPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            //{
            //    Parent = parentPanel,
            //    Type = PanelType.Horizontal,
            //};
            //parentPanel.Children = new[] { childPanel };
            //parentPanel.Render();

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}