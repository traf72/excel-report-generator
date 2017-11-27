using ClosedXML.Excel;
using ExcelReporter.Enums;
using ExcelReporter.Rendering.Panels.ExcelPanels;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourcePanelRenderTests
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

            var parentPanel = new ExcelDataSourcePanel("m:PanelsDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            {
                Type = PanelType.Horizontal,
            };
            var childPanel = new ExcelDataSourcePanel("m:PanelsDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_HorizontalPanels_ChildLeft_Test),
                nameof(Test_HorizontalPanelsGrouping_ChildLeft_ParentCellsShiftChildCellsShift)), ws.Workbook);

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

            var parentPanel = new ExcelDataSourcePanel("m:PanelsDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.Row,
            };
            var childPanel = new ExcelDataSourcePanel("m:PanelsDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_HorizontalPanels_ChildLeft_Test),
                nameof(Test_HorizontalPanelsGrouping_ChildLeft_ParentRowShiftChildCellsShift)), ws.Workbook);

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

            var parentPanel = new ExcelDataSourcePanel("m:PanelsDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.Row,
            };
            var childPanel = new ExcelDataSourcePanel("m:PanelsDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.Row,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_HorizontalPanels_ChildLeft_Test),
                nameof(Test_HorizontalPanelsGrouping_ChildLeft_ParentRowShiftChildRowShift)), ws.Workbook);

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

            var parentPanel = new ExcelDataSourcePanel("m:PanelsDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.NoShift,
            };
            var childPanel = new ExcelDataSourcePanel("m:PanelsDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.Row,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_HorizontalPanels_ChildLeft_Test),
                nameof(Test_HorizontalPanelsGrouping_ChildLeft_ParentNoShiftChildRowShift)), ws.Workbook);

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

            var parentPanel = new ExcelDataSourcePanel("m:PanelsDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            {
                Type = PanelType.Horizontal,
            };
            var childPanel = new ExcelDataSourcePanel("m:PanelsDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_HorizontalPanels_ChildLeft_Test),
                nameof(Test_HorizontalPanelsGrouping_ChildLeft_ParentCellsShiftChildCellsShift_WithFictitiousColumn)), ws.Workbook);

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

            var parentPanel = new ExcelDataSourcePanel("m:PanelsDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.Row,
            };
            var childPanel = new ExcelDataSourcePanel("m:PanelsDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_HorizontalPanels_ChildLeft_Test),
                nameof(Test_HorizontalPanelsGrouping_ChildLeft_ParentRowShiftChildCellsShift_WithFictitiousColumn)), ws.Workbook);

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

            var parentPanel = new ExcelDataSourcePanel("m:PanelsDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.Row,
            };
            var childPanel = new ExcelDataSourcePanel("m:PanelsDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.Row,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_HorizontalPanels_ChildLeft_Test),
                nameof(Test_HorizontalPanelsGrouping_ChildLeft_ParentRowShiftChildRowShift_WithFictitiousColumn)), ws.Workbook);

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

            var parentPanel = new ExcelDataSourcePanel("m:PanelsDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.NoShift,
            };
            var childPanel = new ExcelDataSourcePanel("m:PanelsDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_HorizontalPanels_ChildLeft_Test),
                nameof(Test_HorizontalPanelsGrouping_ChildLeft_ParentNoShiftChildCellsShift_WithFictitiousColumn)), ws.Workbook);

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

            var parentPanel = new ExcelDataSourcePanel("m:PanelsDataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report)
            {
                Type = PanelType.Horizontal,
                AfterRenderMethodName = "AfterRenderParentDataSourcePanelChildLeft",
            };
            var childPanel = new ExcelDataSourcePanel("m:PanelsDataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_HorizontalPanels_ChildLeft_Test),
                nameof(Test_HorizontalPanelsGrouping_ChildLeft_ParentCellsShiftChildCellsShift_WithFictitiousColumnWhichDeleteAfterRender)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}