using ClosedXML.Excel;
using ExcelReporter.Implementations.Panels.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Implementations.Panels.Excel.PanelRenderTests.DataSourcePanelRenderTests
{
    [TestClass]
    public class DataSourcePanelRender_WithGrouping_VerticalPanels_ChildTop_Test
    {
        [TestMethod]
        public void TestVerticalPanelsGroupingChildTop()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 3, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange child = ws.Range(2, 2, 2, 5);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            child.Range(1, 2, 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            child.Range(1, 2, 1, 4).Style.Border.OutsideBorderColor = XLColor.Red;

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(3, 2).Value = "{di:Name}";
            ws.Cell(3, 3).Value = "{di:Date}";

            ws.Cell(2, 3).Value = "{di:Field1}";
            ws.Cell(2, 4).Value = "{di:Field2}";
            ws.Cell(2, 5).Value = "{di:parent:Sum}";

            //ws.Cell(1, 1).Value = "{di:Name}";
            //ws.Cell(4, 1).Value = "{di:Name}";
            //ws.Cell(1, 6).Value = "{di:Name}";
            //ws.Cell(4, 6).Value = "{di:Name}";
            //ws.Cell(3, 1).Value = "{di:Name}";
            //ws.Cell(3, 6).Value = "{di:Name}";
            //ws.Cell(1, 4).Value = "{di:Name}";
            //ws.Cell(4, 4).Value = "{di:Name}";

            var parentPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report);
            var childPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
            };
            parentPanel.Children = new[] { childPanel };
            //parentPanel.Render();

            //Assert.AreEqual(29, ws.CellsUsed().Count());
            //Assert.AreEqual("Test1_01.11.2017", ws.Cell(2, 2).Value);
            //Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(2, 3).Value);
            //Assert.AreEqual(278.8, ws.Cell(2, 4).Value);
            //Assert.AreEqual("15_345", ws.Cell(2, 5).Value);
            //Assert.AreEqual(15d, ws.Cell(3, 2).Value);
            //Assert.AreEqual(345d, ws.Cell(3, 3).Value);
            //Assert.AreEqual("String parameter", ws.Cell(3, 4).Value);
            //Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.TopBorder);
            //Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.BottomBorder);
            //Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.LeftBorder);
            //Assert.AreEqual(XLCellValues.Number, ws.Cell(2, 4).DataType);

            //Assert.AreEqual("Test2_02.11.2017", ws.Cell(4, 2).Value);
            //Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(4, 3).Value);
            //Assert.AreEqual(550d, ws.Cell(4, 4).Value);
            //Assert.AreEqual("76_753465", ws.Cell(4, 5).Value);
            //Assert.AreEqual(76d, ws.Cell(5, 2).Value);
            //Assert.AreEqual(753465d, ws.Cell(5, 3).Value);
            //Assert.AreEqual("String parameter", ws.Cell(5, 4).Value);
            //Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.TopBorder);
            //Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.BottomBorder);
            //Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.LeftBorder);
            //Assert.AreEqual(XLCellValues.Number, ws.Cell(4, 4).DataType);

            //Assert.AreEqual("Test3_03.11.2017", ws.Cell(6, 2).Value);
            //Assert.AreEqual(new DateTime(2017, 11, 3), ws.Cell(6, 3).Value);
            //Assert.AreEqual(27504d, ws.Cell(6, 4).Value);
            //Assert.AreEqual("1533_5456", ws.Cell(6, 5).Value);
            //Assert.AreEqual(1533d, ws.Cell(7, 2).Value);
            //Assert.AreEqual(5456d, ws.Cell(7, 3).Value);
            //Assert.AreEqual("String parameter", ws.Cell(7, 4).Value);
            //Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.TopBorder);
            //Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.BottomBorder);
            //Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.LeftBorder);
            //Assert.AreEqual(XLCellValues.Number, ws.Cell(6, 4).DataType);

            //Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            //Assert.AreEqual("{di:Name}", ws.Cell(4, 1).Value);
            //Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            //Assert.AreEqual("{di:Name}", ws.Cell(4, 6).Value);
            //Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            //Assert.AreEqual("{di:Name}", ws.Cell(3, 6).Value);
            //Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            //Assert.AreEqual("{di:Name}", ws.Cell(8, 4).Value);

            //Assert.AreEqual(0, ws.NamedRanges.Count());
            //Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            //Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            report.Workbook.SaveAs("test4_templ.xlsx");
        }

        [TestMethod]
        public void TestVerticalPanelsGroupingChildTopWithFictitiousRow()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 4, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange child = ws.Range(3, 2, 3, 5);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            child.Range(1, 2, 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            child.Range(1, 2, 1, 4).Style.Border.OutsideBorderColor = XLColor.Red;

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(4, 2).Value = "{di:Name}";
            ws.Cell(4, 3).Value = "{di:Date}";

            ws.Cell(3, 3).Value = "{di:Field1}";
            ws.Cell(3, 4).Value = "{di:Field2}";
            ws.Cell(3, 5).Value = "{di:parent:Sum}";

            //ws.Cell(1, 1).Value = "{di:Name}";
            //ws.Cell(4, 1).Value = "{di:Name}";
            //ws.Cell(1, 6).Value = "{di:Name}";
            //ws.Cell(4, 6).Value = "{di:Name}";
            //ws.Cell(3, 1).Value = "{di:Name}";
            //ws.Cell(3, 6).Value = "{di:Name}";
            //ws.Cell(1, 4).Value = "{di:Name}";
            //ws.Cell(4, 4).Value = "{di:Name}";

            var parentPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            {
                //AfterRenderMethodName = "AfterRenderParentDataSourcePanelChildTop",
            };
            var childPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            //Assert.AreEqual(29, ws.CellsUsed().Count());
            //Assert.AreEqual("Test1_01.11.2017", ws.Cell(2, 2).Value);
            //Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(2, 3).Value);
            //Assert.AreEqual(278.8, ws.Cell(2, 4).Value);
            //Assert.AreEqual("15_345", ws.Cell(2, 5).Value);
            //Assert.AreEqual(15d, ws.Cell(3, 2).Value);
            //Assert.AreEqual(345d, ws.Cell(3, 3).Value);
            //Assert.AreEqual("String parameter", ws.Cell(3, 4).Value);
            //Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.TopBorder);
            //Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.BottomBorder);
            //Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.LeftBorder);
            //Assert.AreEqual(XLCellValues.Number, ws.Cell(2, 4).DataType);

            //Assert.AreEqual("Test2_02.11.2017", ws.Cell(4, 2).Value);
            //Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(4, 3).Value);
            //Assert.AreEqual(550d, ws.Cell(4, 4).Value);
            //Assert.AreEqual("76_753465", ws.Cell(4, 5).Value);
            //Assert.AreEqual(76d, ws.Cell(5, 2).Value);
            //Assert.AreEqual(753465d, ws.Cell(5, 3).Value);
            //Assert.AreEqual("String parameter", ws.Cell(5, 4).Value);
            //Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.TopBorder);
            //Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.BottomBorder);
            //Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.LeftBorder);
            //Assert.AreEqual(XLCellValues.Number, ws.Cell(4, 4).DataType);

            //Assert.AreEqual("Test3_03.11.2017", ws.Cell(6, 2).Value);
            //Assert.AreEqual(new DateTime(2017, 11, 3), ws.Cell(6, 3).Value);
            //Assert.AreEqual(27504d, ws.Cell(6, 4).Value);
            //Assert.AreEqual("1533_5456", ws.Cell(6, 5).Value);
            //Assert.AreEqual(1533d, ws.Cell(7, 2).Value);
            //Assert.AreEqual(5456d, ws.Cell(7, 3).Value);
            //Assert.AreEqual("String parameter", ws.Cell(7, 4).Value);
            //Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.TopBorder);
            //Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.BottomBorder);
            //Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.LeftBorder);
            //Assert.AreEqual(XLCellValues.Number, ws.Cell(6, 4).DataType);

            //Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            //Assert.AreEqual("{di:Name}", ws.Cell(4, 1).Value);
            //Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            //Assert.AreEqual("{di:Name}", ws.Cell(4, 6).Value);
            //Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            //Assert.AreEqual("{di:Name}", ws.Cell(3, 6).Value);
            //Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            //Assert.AreEqual("{di:Name}", ws.Cell(8, 4).Value);

            //Assert.AreEqual(0, ws.NamedRanges.Count());
            //Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            //Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            report.Workbook.SaveAs("test5.xlsx");
        }
    }
}