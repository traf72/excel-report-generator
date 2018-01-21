using ClosedXML.Excel;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourcePanelRenderTests
{
    [TestClass]
    public class DataSourcePanelRender_WithGrouping_VerticalPanels_ChildBottom_Test
    {
        [TestMethod]
        public void Test_VerticalPanelsGrouping_ChildBottom_ParentCellsShiftChildCellsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 3, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange child = ws.Range(3, 2, 3, 5);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            child.Range(1, 2, 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            child.Range(1, 2, 1, 4).Style.Border.OutsideBorderColor = XLColor.Red;

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";

            ws.Cell(3, 3).Value = "{di:Field1}";
            ws.Cell(3, 4).Value = "{di:Field2}";
            ws.Cell(3, 5).Value = "{di:parent:Sum}";

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(4, 1).Value = "{di:Name}";
            ws.Cell(1, 6).Value = "{di:Name}";
            ws.Cell(4, 6).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 6).Value = "{di:Name}";
            ws.Cell(1, 4).Value = "{di:Name}";
            ws.Cell(4, 4).Value = "{di:Name}";

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report, report.TemplateProcessor);
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
            };
            parentPanel.Children = new[] { childPanel };
            IXLRange resultRange = parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 9, 5), resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_VerticalPanels_ChildBottom_Test),
                "ParentCellsShiftChildCellsShift"), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_VerticalPanelsGrouping_ChildBottom_ParentRowShiftChildCellsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 3, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange child = ws.Range(3, 2, 3, 5);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            child.Range(1, 2, 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            child.Range(1, 2, 1, 4).Style.Border.OutsideBorderColor = XLColor.Red;

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";

            ws.Cell(3, 3).Value = "{di:Field1}";
            ws.Cell(3, 4).Value = "{di:Field2}";
            ws.Cell(3, 5).Value = "{di:parent:Sum}";

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(4, 1).Value = "{di:Name}";
            ws.Cell(1, 6).Value = "{di:Name}";
            ws.Cell(4, 6).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 6).Value = "{di:Name}";
            ws.Cell(1, 4).Value = "{di:Name}";
            ws.Cell(4, 4).Value = "{di:Name}";

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report, report.TemplateProcessor)
            {
                ShiftType = ShiftType.Row
            };
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
            };
            parentPanel.Children = new[] { childPanel };
            IXLRange resultRange = parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 9, 5), resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_VerticalPanels_ChildBottom_Test),
                "ParentRowShiftChildCellsShift"), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_VerticalPanelsGrouping_ChildBottom_ParentRowShiftChildRowShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 3, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange child = ws.Range(3, 2, 3, 5);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            child.Range(1, 2, 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            child.Range(1, 2, 1, 4).Style.Border.OutsideBorderColor = XLColor.Red;

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";

            ws.Cell(3, 3).Value = "{di:Field1}";
            ws.Cell(3, 4).Value = "{di:Field2}";
            ws.Cell(3, 5).Value = "{di:parent:Sum}";

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(4, 1).Value = "{di:Name}";
            ws.Cell(1, 6).Value = "{di:Name}";
            ws.Cell(4, 6).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 6).Value = "{di:Name}";
            ws.Cell(1, 4).Value = "{di:Name}";
            ws.Cell(4, 4).Value = "{di:Name}";

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report, report.TemplateProcessor)
            {
                ShiftType = ShiftType.Row
            };
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
                ShiftType = ShiftType.Row
            };
            parentPanel.Children = new[] { childPanel };
            IXLRange resultRange = parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 9, 5), resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_VerticalPanels_ChildBottom_Test),
                "ParentRowShiftChildRowShift"), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_VerticalPanelsGrouping_ChildBottom_ParentNoShiftChildCellsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 3, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange child = ws.Range(3, 2, 3, 5);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            child.Range(1, 2, 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            child.Range(1, 2, 1, 4).Style.Border.OutsideBorderColor = XLColor.Red;

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";

            ws.Cell(3, 3).Value = "{di:Field1}";
            ws.Cell(3, 4).Value = "{di:Field2}";
            ws.Cell(3, 5).Value = "{di:parent:Sum}";

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(4, 1).Value = "{di:Name}";
            ws.Cell(1, 6).Value = "{di:Name}";
            ws.Cell(4, 6).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Name}";
            ws.Cell(3, 6).Value = "{di:Name}";
            ws.Cell(1, 4).Value = "{di:Name}";
            ws.Cell(4, 4).Value = "{di:Name}";

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report, report.TemplateProcessor)
            {
                ShiftType = ShiftType.NoShift
            };
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
            };
            parentPanel.Children = new[] { childPanel };
            IXLRange resultRange = parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 9, 5), resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_VerticalPanels_ChildBottom_Test),
                "ParentNoShiftChildCellsShift"), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_VerticalPanelsGrouping_ChildBottom_ParentCellsShiftChildCellsShift_WithFictitiousRow()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 4, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange child = ws.Range(3, 2, 3, 5);
            child.AddToNamed("ChildRange");

            child.Range(1, 2, 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            child.Range(1, 2, 1, 4).Style.Border.OutsideBorderColor = XLColor.Red;

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";

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

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report, report.TemplateProcessor);
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.Workbook.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
            };
            parentPanel.Children = new[] { childPanel };
            IXLRange resultRange = parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 12, 5), resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_VerticalPanels_ChildBottom_Test),
                "ParentCellsShiftChildCellsShift_WithFictitiousRow"), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_VerticalPanelsGrouping_ChildBottom_ParentRowShiftChildCellsShift_WithFictitiousRow()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 4, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange child = ws.Range(3, 2, 3, 5);
            child.AddToNamed("ChildRange");

            child.Range(1, 2, 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            child.Range(1, 2, 1, 4).Style.Border.OutsideBorderColor = XLColor.Red;

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";

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

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report, report.TemplateProcessor)
            {
                ShiftType = ShiftType.Row,
            };
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.Workbook.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
            };
            parentPanel.Children = new[] { childPanel };
            IXLRange resultRange = parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 12, 5), resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_VerticalPanels_ChildBottom_Test),
                "ParentRowShiftChildCellsShift_WithFictitiousRow"), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_VerticalPanelsGrouping_ChildBottom_ParentRowShiftChildRowShift_WithFictitiousRow()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 4, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange child = ws.Range(3, 2, 3, 5);
            child.AddToNamed("ChildRange");

            child.Range(1, 2, 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            child.Range(1, 2, 1, 4).Style.Border.OutsideBorderColor = XLColor.Red;

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";

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

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report, report.TemplateProcessor)
            {
                ShiftType = ShiftType.Row,
            };
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.Workbook.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
                ShiftType = ShiftType.Row,
            };
            parentPanel.Children = new[] { childPanel };
            IXLRange resultRange = parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 12, 5), resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_VerticalPanels_ChildBottom_Test),
                "ParentRowShiftChildRowShift_WithFictitiousRow"), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_VerticalPanelsGrouping_ChildBottom_ParentNoShiftChildRowShift_WithFictitiousRow()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 4, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange child = ws.Range(3, 2, 3, 5);
            child.AddToNamed("ChildRange");

            child.Range(1, 2, 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            child.Range(1, 2, 1, 4).Style.Border.OutsideBorderColor = XLColor.Red;

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";

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

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report, report.TemplateProcessor)
            {
                ShiftType = ShiftType.NoShift,
            };
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.Workbook.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
                ShiftType = ShiftType.Row,
            };
            parentPanel.Children = new[] { childPanel };
            IXLRange resultRange = parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 12, 5), resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_VerticalPanels_ChildBottom_Test),
                "ParentNoShiftChildRowShift_WithFictitiousRow"), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_VerticalPanelsGrouping_ChildBottom_ParentCellsShiftChildCellsShift_WithFictitiousRowWhichDeleteAfterRender()
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

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";

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

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report, report.TemplateProcessor)
            {
                AfterDataItemRenderMethodName = "AfterRenderParentDataSourcePanelChildBottom",
            };
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
            };
            parentPanel.Children = new[] { childPanel };
            IXLRange resultRange = parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 9, 5), resultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_VerticalPanels_ChildBottom_Test),
                "ParentCellsShiftChildCellsShift_WithFictitiousRowWhichDeleteAfterRender"), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}