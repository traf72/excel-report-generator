using ClosedXML.Excel;
using ExcelReporter.Enums;
using ExcelReporter.Implementations.Panels.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace ExcelReporter.Tests.Implementations.Panels.Excel.PanelRenderTests.DataSourcePanelRenderTests
{
    [TestClass]
    public class DataSourcePanelRender_WithGrouping_VerticalPanels_ChildTop_Test
    {
        [TestMethod]
        public void Test_VerticalPanelsGrouping_ChildTop_ParentCellsShiftChildCellsShift()
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

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(4, 1).Value = "{di:Name}";
            ws.Cell(1, 6).Value = "{di:Name}";
            ws.Cell(4, 6).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 6).Value = "{di:Name}";
            ws.Cell(1, 4).Value = "{di:Name}";
            ws.Cell(4, 4).Value = "{di:Name}";

            var parentPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report);
            var childPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(29, ws.CellsUsed().Count());
            Assert.AreEqual(string.Empty, ws.Cell(2, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child1_F1", ws.Cell(2, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(2, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(2, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child1_F2", ws.Cell(2, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(2, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.LeftBorder);

            Assert.AreEqual(55.76d, ws.Cell(2, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(2, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(3, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F1", ws.Cell(3, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F2", ws.Cell(3, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 4).Style.Border.LeftBorder);

            Assert.AreEqual(55.76d, ws.Cell(3, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F1", ws.Cell(4, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F2", ws.Cell(4, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.LeftBorder);

            Assert.AreEqual(55.76d, ws.Cell(4, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test1", ws.Cell(5, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(5, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test2", ws.Cell(6, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(6, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(6, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(6, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(7, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F1", ws.Cell(7, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(7, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(7, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F2", ws.Cell(7, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(7, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 4).Style.Border.LeftBorder);

            Assert.AreEqual(5500.8d, ws.Cell(7, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(7, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(8, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child2_F1", ws.Cell(8, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child2_F2", ws.Cell(8, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 4).Style.Border.LeftBorder);

            Assert.AreEqual(5500.8d, ws.Cell(8, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test3", ws.Cell(9, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(new DateTime(2017, 11, 3), ws.Cell(9, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(9, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(9, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 5).Style.Border.LeftBorder);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(4, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(10, 4).Value);

            Assert.AreEqual(1, ws.NamedRanges.Count());
            Assert.AreEqual(1, ws.NamedRange("ChildRange").Ranges.Count);
            Assert.AreEqual(ws.Cell(8, 2), ws.NamedRange("ChildRange").Ranges.ElementAt(0).FirstCell());
            Assert.AreEqual(ws.Cell(8, 5), ws.NamedRange("ChildRange").Ranges.ElementAt(0).LastCell());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_VerticalPanelsGrouping_ChildTop_ParentRowShiftChildCellsShift()
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

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(4, 1).Value = "{di:Name}";
            ws.Cell(1, 6).Value = "{di:Name}";
            ws.Cell(4, 6).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 6).Value = "{di:Name}";
            ws.Cell(1, 4).Value = "{di:Name}";
            ws.Cell(4, 4).Value = "{di:Name}";

            var parentPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            {
                ShiftType = ShiftType.Row,
            };
            var childPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(29, ws.CellsUsed().Count());
            Assert.AreEqual(string.Empty, ws.Cell(2, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child1_F1", ws.Cell(2, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(2, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(2, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child1_F2", ws.Cell(2, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(2, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.LeftBorder);

            Assert.AreEqual(55.76d, ws.Cell(2, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(2, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(3, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F1", ws.Cell(3, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F2", ws.Cell(3, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 4).Style.Border.LeftBorder);

            Assert.AreEqual(55.76d, ws.Cell(3, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F1", ws.Cell(4, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F2", ws.Cell(4, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.LeftBorder);

            Assert.AreEqual(55.76d, ws.Cell(4, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test1", ws.Cell(5, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(5, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test2", ws.Cell(6, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(6, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(6, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(6, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(7, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F1", ws.Cell(7, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(7, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(7, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F2", ws.Cell(7, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(7, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 4).Style.Border.LeftBorder);

            Assert.AreEqual(5500.8d, ws.Cell(7, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(7, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(8, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child2_F1", ws.Cell(8, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child2_F2", ws.Cell(8, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 4).Style.Border.LeftBorder);

            Assert.AreEqual(5500.8d, ws.Cell(8, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test3", ws.Cell(9, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(new DateTime(2017, 11, 3), ws.Cell(9, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(9, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(9, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 5).Style.Border.LeftBorder);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(8, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(8, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(5, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(5, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(10, 4).Value);

            Assert.AreEqual(1, ws.NamedRanges.Count());
            Assert.AreEqual(1, ws.NamedRange("ChildRange").Ranges.Count);
            Assert.AreEqual(ws.Cell(8, 2), ws.NamedRange("ChildRange").Ranges.ElementAt(0).FirstCell());
            Assert.AreEqual(ws.Cell(8, 5), ws.NamedRange("ChildRange").Ranges.ElementAt(0).LastCell());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_VerticalPanelsGrouping_ChildTop_ParentRowShiftChildRowShift()
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

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(4, 1).Value = "{di:Name}";
            ws.Cell(1, 6).Value = "{di:Name}";
            ws.Cell(4, 6).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 6).Value = "{di:Name}";
            ws.Cell(1, 4).Value = "{di:Name}";
            ws.Cell(4, 4).Value = "{di:Name}";

            var parentPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            {
                ShiftType = ShiftType.Row,
            };
            var childPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
                ShiftType = ShiftType.Row,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(29, ws.CellsUsed().Count());
            Assert.AreEqual(string.Empty, ws.Cell(2, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child1_F1", ws.Cell(2, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(2, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(2, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child1_F2", ws.Cell(2, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(2, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.LeftBorder);

            Assert.AreEqual(55.76d, ws.Cell(2, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(2, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(3, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F1", ws.Cell(3, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F2", ws.Cell(3, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 4).Style.Border.LeftBorder);

            Assert.AreEqual(55.76d, ws.Cell(3, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F1", ws.Cell(4, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F2", ws.Cell(4, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.LeftBorder);

            Assert.AreEqual(55.76d, ws.Cell(4, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test1", ws.Cell(5, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(5, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test2", ws.Cell(6, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(6, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(6, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(6, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(7, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F1", ws.Cell(7, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(7, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(7, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F2", ws.Cell(7, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(7, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 4).Style.Border.LeftBorder);

            Assert.AreEqual(5500.8d, ws.Cell(7, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(7, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(8, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child2_F1", ws.Cell(8, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child2_F2", ws.Cell(8, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 4).Style.Border.LeftBorder);

            Assert.AreEqual(5500.8d, ws.Cell(8, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test3", ws.Cell(9, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(new DateTime(2017, 11, 3), ws.Cell(9, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(9, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(9, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 5).Style.Border.LeftBorder);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(9, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(9, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(10, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(10, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(10, 4).Value);

            Assert.AreEqual(1, ws.NamedRanges.Count());
            Assert.AreEqual(1, ws.NamedRange("ChildRange").Ranges.Count);
            Assert.AreEqual(ws.Cell(8, 2), ws.NamedRange("ChildRange").Ranges.ElementAt(0).FirstCell());
            Assert.AreEqual(ws.Cell(8, 5), ws.NamedRange("ChildRange").Ranges.ElementAt(0).LastCell());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_VerticalPanelsGrouping_ChildTop_ParentNoShiftChildRowShift()
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

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(4, 1).Value = "{di:Name}";
            ws.Cell(1, 6).Value = "{di:Name}";
            ws.Cell(4, 6).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 6).Value = "{di:Name}";
            ws.Cell(1, 4).Value = "{di:Name}";
            ws.Cell(4, 4).Value = "{di:Name}";

            var parentPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            {
                ShiftType = ShiftType.NoShift,
            };
            var childPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
                ShiftType = ShiftType.Row,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(26, ws.CellsUsed().Count());
            Assert.AreEqual(string.Empty, ws.Cell(2, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child1_F1", ws.Cell(2, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(2, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(2, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child1_F2", ws.Cell(2, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(2, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.LeftBorder);

            Assert.AreEqual(55.76d, ws.Cell(2, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(2, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(3, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F1", ws.Cell(3, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F2", ws.Cell(3, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 4).Style.Border.LeftBorder);

            Assert.AreEqual(55.76d, ws.Cell(3, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F1", ws.Cell(4, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F2", ws.Cell(4, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.LeftBorder);

            Assert.AreEqual(55.76d, ws.Cell(4, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test1", ws.Cell(5, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(5, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test2", ws.Cell(6, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(6, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(6, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(6, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(7, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F1", ws.Cell(7, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(7, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(7, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F2", ws.Cell(7, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(7, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 4).Style.Border.LeftBorder);

            Assert.AreEqual(5500.8d, ws.Cell(7, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(7, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(8, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child2_F1", ws.Cell(8, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child2_F2", ws.Cell(8, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 4).Style.Border.LeftBorder);

            Assert.AreEqual(5500.8d, ws.Cell(8, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test3", ws.Cell(9, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(new DateTime(2017, 11, 3), ws.Cell(9, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(9, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(9, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 5).Style.Border.LeftBorder);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(5, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(5, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);

            Assert.AreEqual(1, ws.NamedRanges.Count());
            Assert.AreEqual(1, ws.NamedRange("ChildRange").Ranges.Count);
            Assert.AreEqual(ws.Cell(4, 2), ws.NamedRange("ChildRange").Ranges.ElementAt(0).FirstCell());
            Assert.AreEqual(ws.Cell(4, 5), ws.NamedRange("ChildRange").Ranges.ElementAt(0).LastCell());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_VerticalPanelsGrouping_ChildTop_ParentCellsShiftChildCellsShift_WithFictitiousRow()
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

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(5, 1).Value = "{di:Name}";
            ws.Cell(1, 6).Value = "{di:Name}";
            ws.Cell(5, 6).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 6).Value = "{di:Name}";
            ws.Cell(1, 4).Value = "{di:Name}";
            ws.Cell(5, 4).Value = "{di:Name}";

            var parentPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report);
            var childPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(29, ws.CellsUsed().Count());
            Assert.AreEqual(string.Empty, ws.Cell(2, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(2, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(2, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(3, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child1_F1", ws.Cell(3, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child1_F2", ws.Cell(3, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 4).Style.Border.LeftBorder);

            Assert.AreEqual(55.76d, ws.Cell(3, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F1", ws.Cell(4, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F2", ws.Cell(4, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.LeftBorder);

            Assert.AreEqual(55.76d, ws.Cell(4, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F1", ws.Cell(5, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F2", ws.Cell(5, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.LeftBorder);

            Assert.AreEqual(55.76d, ws.Cell(5, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test1", ws.Cell(6, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(6, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(6, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(6, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(7, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(7, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(7, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(7, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test2", ws.Cell(8, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(8, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(8, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(8, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(9, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(9, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(9, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(9, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(10, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(10, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(10, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(10, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(10, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(10, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F1", ws.Cell(10, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(10, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(10, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(10, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(10, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(10, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(10, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(10, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F2", ws.Cell(10, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(10, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(10, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(10, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(10, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(10, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(10, 4).Style.Border.LeftBorder);

            Assert.AreEqual(5500.8d, ws.Cell(10, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(10, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(10, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(10, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(10, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(10, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(10, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(10, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(11, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(11, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(11, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(11, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(11, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(11, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child2_F1", ws.Cell(11, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(11, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(11, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(11, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(11, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(11, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(11, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(11, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child2_F2", ws.Cell(11, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(11, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(11, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(11, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(11, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(11, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(11, 4).Style.Border.LeftBorder);

            Assert.AreEqual(5500.8d, ws.Cell(11, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(11, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(11, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(11, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(11, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(11, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(11, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(11, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test3", ws.Cell(12, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(12, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(12, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(12, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(12, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(12, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(12, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(new DateTime(2017, 11, 3), ws.Cell(12, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(12, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(12, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(12, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(12, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(12, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(12, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(12, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(12, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(12, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(12, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(12, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(12, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(12, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(12, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(12, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(12, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(12, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(12, 5).Style.Border.LeftBorder);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(5, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(5, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(13, 4).Value);

            Assert.AreEqual(1, ws.NamedRanges.Count());
            Assert.AreEqual(1, ws.NamedRange("ChildRange").Ranges.Count);
            Assert.AreEqual(ws.Cell(11, 2), ws.NamedRange("ChildRange").Ranges.ElementAt(0).FirstCell());
            Assert.AreEqual(ws.Cell(11, 5), ws.NamedRange("ChildRange").Ranges.ElementAt(0).LastCell());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_VerticalPanelsGrouping_ChildTop_ParentCellsShiftChildRowShift_WithFictitiousRow()
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

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(5, 1).Value = "{di:Name}";
            ws.Cell(1, 6).Value = "{di:Name}";
            ws.Cell(5, 6).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 6).Value = "{di:Name}";
            ws.Cell(1, 4).Value = "{di:Name}";
            ws.Cell(5, 4).Value = "{di:Name}";

            var parentPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report);
            var childPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
                ShiftType = ShiftType.Row,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(29, ws.CellsUsed().Count());
            Assert.AreEqual(string.Empty, ws.Cell(2, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(2, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(2, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(3, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child1_F1", ws.Cell(3, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child1_F2", ws.Cell(3, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 4).Style.Border.LeftBorder);

            Assert.AreEqual(55.76d, ws.Cell(3, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F1", ws.Cell(4, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F2", ws.Cell(4, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.LeftBorder);

            Assert.AreEqual(55.76d, ws.Cell(4, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F1", ws.Cell(5, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F2", ws.Cell(5, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.LeftBorder);

            Assert.AreEqual(55.76d, ws.Cell(5, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test1", ws.Cell(6, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(6, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(6, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(6, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(7, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(7, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(7, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(7, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test2", ws.Cell(8, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(8, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(8, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(8, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(9, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(9, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(9, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(9, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(10, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(10, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(10, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(10, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(10, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(10, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F1", ws.Cell(10, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(10, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(10, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(10, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(10, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(10, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(10, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(10, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F2", ws.Cell(10, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(10, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(10, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(10, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(10, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(10, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(10, 4).Style.Border.LeftBorder);

            Assert.AreEqual(5500.8d, ws.Cell(10, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(10, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(10, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(10, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(10, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(10, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(10, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(10, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(11, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(11, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(11, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(11, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(11, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(11, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child2_F1", ws.Cell(11, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(11, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(11, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(11, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(11, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(11, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(11, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(11, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child2_F2", ws.Cell(11, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(11, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(11, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(11, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(11, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(11, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(11, 4).Style.Border.LeftBorder);

            Assert.AreEqual(5500.8d, ws.Cell(11, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(11, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(11, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(11, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(11, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(11, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(11, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(11, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test3", ws.Cell(12, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(12, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(12, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(12, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(12, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(12, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(12, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(new DateTime(2017, 11, 3), ws.Cell(12, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(12, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(12, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(12, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(12, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(12, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(12, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(12, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(12, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(12, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(12, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(12, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(12, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(12, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(12, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(12, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(12, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(12, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(12, 5).Style.Border.LeftBorder);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(7, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(7, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(5, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(5, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(13, 4).Value);

            Assert.AreEqual(1, ws.NamedRanges.Count());
            Assert.AreEqual(1, ws.NamedRange("ChildRange").Ranges.Count);
            Assert.AreEqual(ws.Cell(11, 2), ws.NamedRange("ChildRange").Ranges.ElementAt(0).FirstCell());
            Assert.AreEqual(ws.Cell(11, 5), ws.NamedRange("ChildRange").Ranges.ElementAt(0).LastCell());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_VerticalPanelsGrouping_ChildTop_ParentNoShiftChildCellsShift_WithFictitiousRowWhichDeleteAfterRender()
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

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(5, 1).Value = "{di:Name}";
            ws.Cell(1, 6).Value = "{di:Name}";
            ws.Cell(5, 6).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 6).Value = "{di:Name}";
            ws.Cell(1, 4).Value = "{di:Name}";
            ws.Cell(5, 4).Value = "{di:Name}";

            var parentPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            {
                AfterRenderMethodName = "AfterRenderParentDataSourcePanelChildTop",
                ShiftType = ShiftType.NoShift,
            };
            var childPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(28, ws.CellsUsed().Count());
            Assert.AreEqual(string.Empty, ws.Cell(2, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child1_F1", ws.Cell(2, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(2, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(2, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child1_F2", ws.Cell(2, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(2, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.LeftBorder);

            Assert.AreEqual(55.76d, ws.Cell(2, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(2, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(3, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F1", ws.Cell(3, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F2", ws.Cell(3, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 4).Style.Border.LeftBorder);

            Assert.AreEqual(55.76d, ws.Cell(3, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F1", ws.Cell(4, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F2", ws.Cell(4, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.LeftBorder);

            Assert.AreEqual(55.76d, ws.Cell(4, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test1", ws.Cell(5, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(5, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test2", ws.Cell(6, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(6, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(6, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(6, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(6, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(6, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(6, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(7, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F1", ws.Cell(7, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(7, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(7, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F2", ws.Cell(7, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(7, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 4).Style.Border.LeftBorder);

            Assert.AreEqual(5500.8d, ws.Cell(7, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(7, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(7, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(7, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(7, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(8, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child2_F1", ws.Cell(8, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child2_F2", ws.Cell(8, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 4).Style.Border.LeftBorder);

            Assert.AreEqual(5500.8d, ws.Cell(8, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(8, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(8, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(8, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(8, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test3", ws.Cell(9, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(new DateTime(2017, 11, 3), ws.Cell(9, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 3).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(9, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 4).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(9, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(9, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(9, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(9, 5).Style.Border.LeftBorder);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(5, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(5, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);

            Assert.AreEqual(1, ws.NamedRanges.Count());
            Assert.AreEqual(1, ws.NamedRange("ChildRange").Ranges.Count);
            Assert.AreEqual(ws.Cell(4, 2), ws.NamedRange("ChildRange").Ranges.ElementAt(0).FirstCell());
            Assert.AreEqual(ws.Cell(4, 5), ws.NamedRange("ChildRange").Ranges.ElementAt(0).LastCell());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}