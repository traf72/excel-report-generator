using ClosedXML.Excel;
using ExcelReporter.Enums;
using ExcelReporter.Implementations.Panels.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace ExcelReporter.Tests.Implementations.Panels.Excel.PanelRenderTests.DataSourcePanelRenderTests
{
    [TestClass]
    public class DataSourcePanelRender_WithGrouping_HorizontalPanels_ChildLeft_Test
    {
        [TestMethod]
        public void Test_HorizontalPanelsGrouping_ChildLeft_ParentCellsShiftChildCellsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 5, 3);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange child = ws.Range(2, 2, 5, 2);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            child.Range(2, 1, 4, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            child.Range(2, 1, 4, 1).Style.Border.OutsideBorderColor = XLColor.Red;

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(2, 3).Value = "{di:Name}";
            ws.Cell(3, 3).Value = "{di:Date}";

            ws.Cell(3, 2).Value = "{di:Field1}";
            ws.Cell(4, 2).Value = "{di:Field2}";
            ws.Cell(5, 2).Value = "{di:parent:Sum}";

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(1, 3).Value = "{di:Name}";
            ws.Cell(1, 4).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 4).Value = "{di:Name}";
            ws.Cell(6, 1).Value = "{di:Name}";
            ws.Cell(6, 3).Value = "{di:Name}";
            ws.Cell(6, 4).Value = "{di:Name}";

            var parentPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            {
                Type = PanelType.Horizontal,
            };
            var childPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
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

            Assert.AreEqual("Test1_Child1_F1", ws.Cell(3, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 2).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child1_F2", ws.Cell(4, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 2).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 2).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F1", ws.Cell(3, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F2", ws.Cell(4, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F1", ws.Cell(3, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F2", ws.Cell(4, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1", ws.Cell(2, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.LeftBorder);

            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(3, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test2", ws.Cell(2, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 6).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 6).Style.Border.LeftBorder);

            Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(3, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 6).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(2, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 7).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F1", ws.Cell(3, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 7).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 7).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F2", ws.Cell(4, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 7).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual(5500.8d, ws.Cell(5, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 7).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 7).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 8).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 8).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 8).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child2_F1", ws.Cell(3, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 8).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 8).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 8).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child2_F2", ws.Cell(4, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 8).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 8).Style.Border.LeftBorderColor);

            Assert.AreEqual(5500.8d, ws.Cell(5, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 8).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 8).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 8).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3", ws.Cell(2, 9).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 9).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 9).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 9).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 9).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 9).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 9).Style.Border.LeftBorder);

            Assert.AreEqual(new DateTime(2017, 11, 3), ws.Cell(3, 9).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 9).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 9).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 9).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 9).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 9).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 9).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 9).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 9).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 9).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 9).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 9).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 9).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 9).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 9).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 9).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 9).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 9).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 9).Style.Border.LeftBorder);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 3).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 4).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 10).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 3).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 4).Value);

            Assert.AreEqual(1, ws.NamedRanges.Count());
            Assert.AreEqual(1, ws.NamedRange("ChildRange").Ranges.Count);
            Assert.AreEqual(ws.Cell(2, 8), ws.NamedRange("ChildRange").Ranges.ElementAt(0).FirstCell());
            Assert.AreEqual(ws.Cell(5, 8), ws.NamedRange("ChildRange").Ranges.ElementAt(0).LastCell());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_HorizontalPanelsGrouping_ChildLeft_ParentRowShiftChildCellsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 5, 3);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange child = ws.Range(2, 2, 5, 2);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            child.Range(2, 1, 4, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            child.Range(2, 1, 4, 1).Style.Border.OutsideBorderColor = XLColor.Red;

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(2, 3).Value = "{di:Name}";
            ws.Cell(3, 3).Value = "{di:Date}";

            ws.Cell(3, 2).Value = "{di:Field1}";
            ws.Cell(4, 2).Value = "{di:Field2}";
            ws.Cell(5, 2).Value = "{di:parent:Sum}";

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(1, 3).Value = "{di:Name}";
            ws.Cell(1, 4).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 4).Value = "{di:Name}";
            ws.Cell(6, 1).Value = "{di:Name}";
            ws.Cell(6, 3).Value = "{di:Name}";
            ws.Cell(6, 4).Value = "{di:Name}";

            var parentPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.Row,
            };
            var childPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
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

            Assert.AreEqual("Test1_Child1_F1", ws.Cell(3, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 2).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child1_F2", ws.Cell(4, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 2).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 2).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F1", ws.Cell(3, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F2", ws.Cell(4, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F1", ws.Cell(3, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F2", ws.Cell(4, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1", ws.Cell(2, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.LeftBorder);

            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(3, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test2", ws.Cell(2, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 6).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 6).Style.Border.LeftBorder);

            Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(3, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 6).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(2, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 7).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F1", ws.Cell(3, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 7).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 7).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F2", ws.Cell(4, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 7).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual(5500.8d, ws.Cell(5, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 7).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 7).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 8).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 8).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 8).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child2_F1", ws.Cell(3, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 8).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 8).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 8).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child2_F2", ws.Cell(4, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 8).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 8).Style.Border.LeftBorderColor);

            Assert.AreEqual(5500.8d, ws.Cell(5, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 8).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 8).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 8).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3", ws.Cell(2, 9).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 9).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 9).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 9).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 9).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 9).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 9).Style.Border.LeftBorder);

            Assert.AreEqual(new DateTime(2017, 11, 3), ws.Cell(3, 9).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 9).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 9).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 9).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 9).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 9).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 9).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 9).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 9).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 9).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 9).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 9).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 9).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 9).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 9).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 9).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 9).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 9).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 9).Style.Border.LeftBorder);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 5).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 8).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 10).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 5).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 8).Value);

            Assert.AreEqual(1, ws.NamedRanges.Count());
            Assert.AreEqual(1, ws.NamedRange("ChildRange").Ranges.Count);
            Assert.AreEqual(ws.Cell(2, 8), ws.NamedRange("ChildRange").Ranges.ElementAt(0).FirstCell());
            Assert.AreEqual(ws.Cell(5, 8), ws.NamedRange("ChildRange").Ranges.ElementAt(0).LastCell());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_HorizontalPanelsGrouping_ChildLeft_ParentRowShiftChildRowShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 5, 3);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange child = ws.Range(2, 2, 5, 2);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            child.Range(2, 1, 4, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            child.Range(2, 1, 4, 1).Style.Border.OutsideBorderColor = XLColor.Red;

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(2, 3).Value = "{di:Name}";
            ws.Cell(3, 3).Value = "{di:Date}";

            ws.Cell(3, 2).Value = "{di:Field1}";
            ws.Cell(4, 2).Value = "{di:Field2}";
            ws.Cell(5, 2).Value = "{di:parent:Sum}";

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(1, 3).Value = "{di:Name}";
            ws.Cell(1, 4).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 4).Value = "{di:Name}";
            ws.Cell(6, 1).Value = "{di:Name}";
            ws.Cell(6, 3).Value = "{di:Name}";
            ws.Cell(6, 4).Value = "{di:Name}";

            var parentPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.Row,
            };
            var childPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
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

            Assert.AreEqual("Test1_Child1_F1", ws.Cell(3, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 2).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child1_F2", ws.Cell(4, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 2).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 2).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F1", ws.Cell(3, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F2", ws.Cell(4, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F1", ws.Cell(3, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F2", ws.Cell(4, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1", ws.Cell(2, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.LeftBorder);

            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(3, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test2", ws.Cell(2, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 6).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 6).Style.Border.LeftBorder);

            Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(3, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 6).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(2, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 7).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F1", ws.Cell(3, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 7).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 7).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F2", ws.Cell(4, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 7).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual(5500.8d, ws.Cell(5, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 7).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 7).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 8).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 8).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 8).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child2_F1", ws.Cell(3, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 8).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 8).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 8).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child2_F2", ws.Cell(4, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 8).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 8).Style.Border.LeftBorderColor);

            Assert.AreEqual(5500.8d, ws.Cell(5, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 8).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 8).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 8).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3", ws.Cell(2, 9).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 9).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 9).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 9).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 9).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 9).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 9).Style.Border.LeftBorder);

            Assert.AreEqual(new DateTime(2017, 11, 3), ws.Cell(3, 9).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 9).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 9).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 9).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 9).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 9).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 9).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 9).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 9).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 9).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 9).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 9).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 9).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 9).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 9).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 9).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 9).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 9).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 9).Style.Border.LeftBorder);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 9).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 10).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 10).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 9).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 10).Value);

            Assert.AreEqual(1, ws.NamedRanges.Count());
            Assert.AreEqual(1, ws.NamedRange("ChildRange").Ranges.Count);
            Assert.AreEqual(ws.Cell(2, 8), ws.NamedRange("ChildRange").Ranges.ElementAt(0).FirstCell());
            Assert.AreEqual(ws.Cell(5, 8), ws.NamedRange("ChildRange").Ranges.ElementAt(0).LastCell());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_HorizontalPanelsGrouping_ChildLeft_ParentNoShiftChildRowShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 5, 3);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange child = ws.Range(2, 2, 5, 2);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            child.Range(2, 1, 4, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            child.Range(2, 1, 4, 1).Style.Border.OutsideBorderColor = XLColor.Red;

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(2, 3).Value = "{di:Name}";
            ws.Cell(3, 3).Value = "{di:Date}";

            ws.Cell(3, 2).Value = "{di:Field1}";
            ws.Cell(4, 2).Value = "{di:Field2}";
            ws.Cell(5, 2).Value = "{di:parent:Sum}";

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(1, 3).Value = "{di:Name}";
            ws.Cell(1, 4).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 4).Value = "{di:Name}";
            ws.Cell(6, 1).Value = "{di:Name}";
            ws.Cell(6, 3).Value = "{di:Name}";
            ws.Cell(6, 4).Value = "{di:Name}";

            var parentPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.NoShift,
            };
            var childPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
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

            Assert.AreEqual("Test1_Child1_F1", ws.Cell(3, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 2).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child1_F2", ws.Cell(4, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 2).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 2).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F1", ws.Cell(3, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F2", ws.Cell(4, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F1", ws.Cell(3, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F2", ws.Cell(4, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1", ws.Cell(2, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.LeftBorder);

            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(3, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test2", ws.Cell(2, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 6).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 6).Style.Border.LeftBorder);

            Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(3, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 6).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(2, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 7).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F1", ws.Cell(3, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 7).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 7).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F2", ws.Cell(4, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 7).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual(5500.8d, ws.Cell(5, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 7).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 7).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 8).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 8).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 8).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child2_F1", ws.Cell(3, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 8).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 8).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 8).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child2_F2", ws.Cell(4, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 8).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 8).Style.Border.LeftBorderColor);

            Assert.AreEqual(5500.8d, ws.Cell(5, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 8).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 8).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 8).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3", ws.Cell(2, 9).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 9).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 9).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 9).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 9).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 9).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 9).Style.Border.LeftBorder);

            Assert.AreEqual(new DateTime(2017, 11, 3), ws.Cell(3, 9).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 9).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 9).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 9).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 9).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 9).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 9).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 9).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 9).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 9).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 9).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 9).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 9).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 9).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 9).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 9).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 9).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 9).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 9).Style.Border.LeftBorder);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 5).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 5).Value);

            Assert.AreEqual(1, ws.NamedRanges.Count());
            Assert.AreEqual(1, ws.NamedRange("ChildRange").Ranges.Count);
            Assert.AreEqual(ws.Cell(2, 4), ws.NamedRange("ChildRange").Ranges.ElementAt(0).FirstCell());
            Assert.AreEqual(ws.Cell(5, 4), ws.NamedRange("ChildRange").Ranges.ElementAt(0).LastCell());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_HorizontalPanelsGrouping_ChildLeft_ParentCellsShiftChildCellsShift_WithFictitiousColumn()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 5, 4);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange child = ws.Range(2, 3, 5, 3);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            child.Range(2, 1, 4, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            child.Range(2, 1, 4, 1).Style.Border.OutsideBorderColor = XLColor.Red;

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(2, 4).Value = "{di:Name}";
            ws.Cell(3, 4).Value = "{di:Date}";

            ws.Cell(3, 3).Value = "{di:Field1}";
            ws.Cell(4, 3).Value = "{di:Field2}";
            ws.Cell(5, 3).Value = "{di:parent:Sum}";

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(1, 3).Value = "{di:Name}";
            ws.Cell(1, 5).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 5).Value = "{di:Name}";
            ws.Cell(6, 1).Value = "{di:Name}";
            ws.Cell(6, 3).Value = "{di:Name}";
            ws.Cell(6, 5).Value = "{di:Name}";

            var parentPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            {
                Type = PanelType.Horizontal,
            };
            var childPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
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

            Assert.AreEqual(string.Empty, ws.Cell(3, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(4, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(5, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.LeftBorder);

            Assert.AreEqual("Test1_Child1_F1", ws.Cell(3, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child1_F2", ws.Cell(4, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.LeftBorder);

            Assert.AreEqual("Test1_Child2_F1", ws.Cell(3, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F2", ws.Cell(4, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test1_Child3_F1", ws.Cell(3, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 5).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F2", ws.Cell(4, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 5).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 5).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1", ws.Cell(2, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 6).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 6).Style.Border.LeftBorder);

            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(3, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 6).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(2, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 7).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(3, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(4, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(5, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 7).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test2", ws.Cell(2, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 8).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 8).Style.Border.LeftBorder);

            Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(3, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 8).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 8).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 8).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 8).Style.Border.LeftBorder);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 3).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 5).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 13).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 3).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 5).Value);

            Assert.AreEqual(1, ws.NamedRanges.Count());
            Assert.AreEqual(1, ws.NamedRange("ChildRange").Ranges.Count);
            Assert.AreEqual(ws.Cell(2, 11), ws.NamedRange("ChildRange").Ranges.ElementAt(0).FirstCell());
            Assert.AreEqual(ws.Cell(5, 11), ws.NamedRange("ChildRange").Ranges.ElementAt(0).LastCell());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_HorizontalPanelsGrouping_ChildLeft_ParentRowShiftChildCellsShift_WithFictitiousColumn()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 5, 4);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange child = ws.Range(2, 3, 5, 3);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            child.Range(2, 1, 4, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            child.Range(2, 1, 4, 1).Style.Border.OutsideBorderColor = XLColor.Red;

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(2, 4).Value = "{di:Name}";
            ws.Cell(3, 4).Value = "{di:Date}";

            ws.Cell(3, 3).Value = "{di:Field1}";
            ws.Cell(4, 3).Value = "{di:Field2}";
            ws.Cell(5, 3).Value = "{di:parent:Sum}";

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(1, 3).Value = "{di:Name}";
            ws.Cell(1, 5).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 5).Value = "{di:Name}";
            ws.Cell(6, 1).Value = "{di:Name}";
            ws.Cell(6, 3).Value = "{di:Name}";
            ws.Cell(6, 5).Value = "{di:Name}";

            var parentPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.Row,
            };
            var childPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
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

            Assert.AreEqual(string.Empty, ws.Cell(3, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(4, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(5, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.LeftBorder);

            Assert.AreEqual("Test1_Child1_F1", ws.Cell(3, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child1_F2", ws.Cell(4, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.LeftBorder);

            Assert.AreEqual("Test1_Child2_F1", ws.Cell(3, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F2", ws.Cell(4, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test1_Child3_F1", ws.Cell(3, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 5).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F2", ws.Cell(4, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 5).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 5).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1", ws.Cell(2, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 6).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 6).Style.Border.LeftBorder);

            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(3, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 6).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(2, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 7).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(3, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(4, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(5, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 7).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test2", ws.Cell(2, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 8).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 8).Style.Border.LeftBorder);

            Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(3, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 8).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 8).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 8).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 8).Style.Border.LeftBorder);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 11).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 13).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 6).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 11).Value);

            Assert.AreEqual(1, ws.NamedRanges.Count());
            Assert.AreEqual(1, ws.NamedRange("ChildRange").Ranges.Count);
            Assert.AreEqual(ws.Cell(2, 11), ws.NamedRange("ChildRange").Ranges.ElementAt(0).FirstCell());
            Assert.AreEqual(ws.Cell(5, 11), ws.NamedRange("ChildRange").Ranges.ElementAt(0).LastCell());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_HorizontalPanelsGrouping_ChildLeft_ParentRowShiftChildRowShift_WithFictitiousColumn()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 5, 4);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange child = ws.Range(2, 3, 5, 3);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            child.Range(2, 1, 4, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            child.Range(2, 1, 4, 1).Style.Border.OutsideBorderColor = XLColor.Red;

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(2, 4).Value = "{di:Name}";
            ws.Cell(3, 4).Value = "{di:Date}";

            ws.Cell(3, 3).Value = "{di:Field1}";
            ws.Cell(4, 3).Value = "{di:Field2}";
            ws.Cell(5, 3).Value = "{di:parent:Sum}";

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(1, 3).Value = "{di:Name}";
            ws.Cell(1, 5).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 5).Value = "{di:Name}";
            ws.Cell(6, 1).Value = "{di:Name}";
            ws.Cell(6, 3).Value = "{di:Name}";
            ws.Cell(6, 5).Value = "{di:Name}";

            var parentPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.Row,
            };
            var childPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
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

            Assert.AreEqual(string.Empty, ws.Cell(3, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(4, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(5, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.LeftBorder);

            Assert.AreEqual("Test1_Child1_F1", ws.Cell(3, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child1_F2", ws.Cell(4, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.LeftBorder);

            Assert.AreEqual("Test1_Child2_F1", ws.Cell(3, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F2", ws.Cell(4, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test1_Child3_F1", ws.Cell(3, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 5).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F2", ws.Cell(4, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 5).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 5).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1", ws.Cell(2, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 6).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 6).Style.Border.LeftBorder);

            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(3, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 6).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(2, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 7).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(3, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(4, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(5, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 7).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test2", ws.Cell(2, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 8).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 8).Style.Border.LeftBorder);

            Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(3, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 8).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 8).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 8).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 8).Style.Border.LeftBorder);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 11).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 13).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 13).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 11).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 13).Value);

            Assert.AreEqual(1, ws.NamedRanges.Count());
            Assert.AreEqual(1, ws.NamedRange("ChildRange").Ranges.Count);
            Assert.AreEqual(ws.Cell(2, 11), ws.NamedRange("ChildRange").Ranges.ElementAt(0).FirstCell());
            Assert.AreEqual(ws.Cell(5, 11), ws.NamedRange("ChildRange").Ranges.ElementAt(0).LastCell());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_HorizontalPanelsGrouping_ChildLeft_ParentNoShiftChildCellsShift_WithFictitiousColumn()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 5, 4);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange child = ws.Range(2, 3, 5, 3);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            child.Range(2, 1, 4, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            child.Range(2, 1, 4, 1).Style.Border.OutsideBorderColor = XLColor.Red;

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(2, 4).Value = "{di:Name}";
            ws.Cell(3, 4).Value = "{di:Date}";

            ws.Cell(3, 3).Value = "{di:Field1}";
            ws.Cell(4, 3).Value = "{di:Field2}";
            ws.Cell(5, 3).Value = "{di:parent:Sum}";

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(1, 3).Value = "{di:Name}";
            ws.Cell(1, 5).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 5).Value = "{di:Name}";
            ws.Cell(6, 1).Value = "{di:Name}";
            ws.Cell(6, 3).Value = "{di:Name}";
            ws.Cell(6, 5).Value = "{di:Name}";

            var parentPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.NoShift,
            };
            var childPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
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

            Assert.AreEqual(string.Empty, ws.Cell(3, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(4, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(5, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.LeftBorder);

            Assert.AreEqual("Test1_Child1_F1", ws.Cell(3, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child1_F2", ws.Cell(4, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.LeftBorder);

            Assert.AreEqual("Test1_Child2_F1", ws.Cell(3, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F2", ws.Cell(4, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test1_Child3_F1", ws.Cell(3, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 5).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F2", ws.Cell(4, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 5).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 5).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1", ws.Cell(2, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 6).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 6).Style.Border.LeftBorder);

            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(3, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 6).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 6).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(2, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 7).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(3, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(4, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(5, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 7).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test2", ws.Cell(2, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 8).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 8).Style.Border.LeftBorder);

            Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(3, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 8).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 8).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 8).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 8).Style.Border.LeftBorder);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 3).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 5).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 3).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 5).Value);

            Assert.AreEqual(1, ws.NamedRanges.Count());
            Assert.AreEqual(1, ws.NamedRange("ChildRange").Ranges.Count);
            Assert.AreEqual(ws.Cell(2, 5), ws.NamedRange("ChildRange").Ranges.ElementAt(0).FirstCell());
            Assert.AreEqual(ws.Cell(5, 5), ws.NamedRange("ChildRange").Ranges.ElementAt(0).LastCell());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_HorizontalPanelsGrouping_ChildLeft_ParentCellsShiftChildCellsShift_WithFictitiousColumnWhichDeleteAfterRender()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 5, 4);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange child = ws.Range(2, 3, 5, 3);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            child.Range(2, 1, 4, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            child.Range(2, 1, 4, 1).Style.Border.OutsideBorderColor = XLColor.Red;

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(2, 4).Value = "{di:Name}";
            ws.Cell(3, 4).Value = "{di:Date}";

            ws.Cell(3, 3).Value = "{di:Field1}";
            ws.Cell(4, 3).Value = "{di:Field2}";
            ws.Cell(5, 3).Value = "{di:parent:Sum}";

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(1, 3).Value = "{di:Name}";
            ws.Cell(1, 5).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 5).Value = "{di:Name}";
            ws.Cell(6, 1).Value = "{di:Name}";
            ws.Cell(6, 3).Value = "{di:Name}";
            ws.Cell(6, 5).Value = "{di:Name}";

            var parentPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            {
                Type = PanelType.Horizontal,
                AfterRenderMethodName = "AfterRenderParentDataSourcePanelChildLeft",
            };
            var childPanel = new ExcelDataSourcePanel("m:TestDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
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

            Assert.AreEqual("Test1_Child1_F1", ws.Cell(3, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 2).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 2).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child1_F2", ws.Cell(4, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 2).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 2).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 2).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 2).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 2).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 2).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 3).Style.Border.LeftBorder);

            Assert.AreEqual("Test1_Child2_F1", ws.Cell(3, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child2_F2", ws.Cell(4, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 3).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 3).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 3).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 3).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 3).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 3).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 4).Style.Border.LeftBorder);

            Assert.AreEqual("Test1_Child3_F1", ws.Cell(3, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1_Child3_F2", ws.Cell(4, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual(55.76d, ws.Cell(5, 4).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 4).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 4).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 4).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 4).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 4).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test1", ws.Cell(2, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 5).Style.Border.LeftBorder);

            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(3, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 5).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 5).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 5).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 5).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 5).Style.Border.LeftBorder);

            Assert.AreEqual("Test2", ws.Cell(2, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 6).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 6).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 6).Style.Border.LeftBorderColor);

            Assert.AreEqual(new DateTime(2017, 11, 2), ws.Cell(3, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 6).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 6).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(4, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 6).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 6).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(5, 6).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 6).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 6).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 6).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 6).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 6).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 6).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 6).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 7).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F1", ws.Cell(3, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 7).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 7).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child1_F2", ws.Cell(4, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 7).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual(5500.8d, ws.Cell(5, 7).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 7).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 7).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 7).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 7).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 7).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 7).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 7).Style.Border.LeftBorderColor);

            Assert.AreEqual(string.Empty, ws.Cell(2, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 8).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 8).Style.Border.LeftBorder);

            Assert.AreEqual("Test3_Child2_F1", ws.Cell(3, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 8).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 8).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(3, 8).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3_Child2_F2", ws.Cell(4, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 8).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(4, 8).Style.Border.LeftBorderColor);

            Assert.AreEqual(5500.8d, ws.Cell(5, 8).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 8).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 8).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 8).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 8).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 8).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 8).Style.Border.LeftBorder);
            Assert.AreEqual(XLColor.Red, ws.Cell(5, 8).Style.Border.LeftBorderColor);

            Assert.AreEqual("Test3", ws.Cell(2, 9).Value);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 9).Style.Border.TopBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 9).Style.Border.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(2, 9).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(2, 9).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 9).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(2, 9).Style.Border.LeftBorder);

            Assert.AreEqual(new DateTime(2017, 11, 3), ws.Cell(3, 9).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 9).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(3, 9).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(3, 9).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 9).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(3, 9).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(4, 9).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 9).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(4, 9).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(4, 9).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 9).Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(4, 9).Style.Border.LeftBorder);

            Assert.AreEqual(string.Empty, ws.Cell(5, 9).Value);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 9).Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 9).Style.Border.RightBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 9).Style.Border.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell(5, 9).Style.Border.BottomBorder);
            Assert.AreEqual(XLColor.Black, ws.Cell(5, 9).Style.Border.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.None, ws.Cell(5, 9).Style.Border.LeftBorder);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 3).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(1, 5).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(3, 10).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 1).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 3).Value);
            Assert.AreEqual("{di:Name}", ws.Cell(6, 5).Value);

            Assert.AreEqual(1, ws.NamedRanges.Count());
            Assert.AreEqual(1, ws.NamedRange("ChildRange").Ranges.Count);
            Assert.AreEqual(ws.Cell(2, 8), ws.NamedRange("ChildRange").Ranges.ElementAt(0).FirstCell());
            Assert.AreEqual(ws.Cell(5, 8), ws.NamedRange("ChildRange").Ranges.ElementAt(0).LastCell());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}