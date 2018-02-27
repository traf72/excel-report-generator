using ClosedXML.Excel;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourcePanelRenderTests
{
    [TestClass]
    public class DataSourcePanelRender_WithGrouping_PageBreaks_Test
    {
        [TestMethod]
        public void Test_HorizontalPageBreaks_WithSimplePanel_CellsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 4, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            IXLRange child1 = ws.Range(3, 2, 3, 5);
            child1.AddToNamed("ChildRange1", XLScope.Worksheet);

            IXLRange child2 = ws.Range(4, 2, 4, 5);
            child2.AddToNamed("ChildRange2", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";

            ws.Cell(3, 3).Value = "{di:Field1}";
            ws.Cell(3, 4).Value = "{di:Field2}";
            ws.Cell(3, 5).Value = "{di:parent:Sum}";

            ws.Cell(4, 2).Value = "{HorizPageBreak}";

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report, report.TemplateProcessor);
            var childPanel1 = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange1"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
            };
            var childPanel2 = new ExcelNamedPanel(ws.NamedRange("ChildRange2"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
            };
            parentPanel.Children = new[] { childPanel1, childPanel2 };
            parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 12, 5), parentPanel.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_PageBreaks_Test),
                "Test_HorizontalPageBreaks"), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_HorizontalPageBreaks_WithoutSimplePanel_CellsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 4, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            IXLRange child = ws.Range(3, 2, 3, 5);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";

            ws.Cell(3, 3).Value = "{di:Field1}";
            ws.Cell(3, 4).Value = "{di:Field2}";
            ws.Cell(3, 5).Value = "{di:parent:Sum}";

            ws.Cell(4, 2).Value = "{ HorizPageBreak }";

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report, report.TemplateProcessor);
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 12, 5), parentPanel.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_PageBreaks_Test),
                "Test_HorizontalPageBreaks"), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_HorizontalPageBreaks_WithSimplePanel_RowShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 4, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            IXLRange child1 = ws.Range(3, 2, 3, 5);
            child1.AddToNamed("ChildRange1", XLScope.Worksheet);

            IXLRange child2 = ws.Range(4, 2, 4, 5);
            child2.AddToNamed("ChildRange2", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";

            ws.Cell(3, 3).Value = "{di:Field1}";
            ws.Cell(3, 4).Value = "{di:Field2}";
            ws.Cell(3, 5).Value = "{di:parent:Sum}";

            ws.Cell(4, 2).Value = "  {  horizpagebreak  }  ";

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report, report.TemplateProcessor);
            var childPanel1 = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange1"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
            };
            var childPanel2 = new ExcelNamedPanel(ws.NamedRange("ChildRange2"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
                ShiftType = ShiftType.Row,
            };
            parentPanel.Children = new[] { childPanel1, childPanel2 };
            parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 12, 5), parentPanel.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_PageBreaks_Test),
                "Test_HorizontalPageBreaks"), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_HorizontalPageBreaks_WithoutSimplePanel_RowShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 4, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            IXLRange child = ws.Range(3, 2, 3, 5);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";

            ws.Cell(3, 3).Value = "{di:Field1}";
            ws.Cell(3, 4).Value = "{di:Field2}";
            ws.Cell(3, 5).Value = "{di:parent:Sum}";

            ws.Cell(4, 2).Value = "{HorizPageBreak}";

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report, report.TemplateProcessor)
            {
                ShiftType = ShiftType.Row,
            };
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
                ShiftType = ShiftType.Row,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 12, 5), parentPanel.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_PageBreaks_Test),
                "Test_HorizontalPageBreaks"), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_VerticalPageBreaks_WithSimplePanel_CellsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 5, 4);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            IXLRange child1 = ws.Range(2, 3, 5, 3);
            child1.AddToNamed("ChildRange1", XLScope.Worksheet);

            IXLRange child2 = ws.Range(2, 4, 5, 4);
            child2.AddToNamed("ChildRange2", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(3, 2).Value = "{di:Date}";

            ws.Cell(3, 3).Value = "{di:Field1}";
            ws.Cell(4, 3).Value = "{di:Field2}";
            ws.Cell(5, 3).Value = "{di:parent:Sum}";

            ws.Cell(2, 4).Value = " {vertPageBreak} ";

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report, report.TemplateProcessor)
            {
                Type = PanelType.Horizontal,
            };
            var childPanel1 = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange1"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
            };
            var childPanel2 = new ExcelNamedPanel(ws.NamedRange("ChildRange2"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
            };
            parentPanel.Children = new[] { childPanel1, childPanel2 };
            parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 5, 12), parentPanel.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_PageBreaks_Test),
                "Test_VerticalPageBreaks"), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_VerticalPageBreaks_WithoutSimplePanel_CellsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 5, 4);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            IXLRange child = ws.Range(2, 3, 5, 3);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(3, 2).Value = "{di:Date}";

            ws.Cell(3, 3).Value = "{di:Field1}";
            ws.Cell(4, 3).Value = "{di:Field2}";
            ws.Cell(5, 3).Value = "{di:parent:Sum}";

            ws.Cell(2, 4).Value = " { VertPageBreak } ";

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report, report.TemplateProcessor)
            {
                Type = PanelType.Horizontal,
            };
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 5, 12), parentPanel.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_PageBreaks_Test),
                "Test_VerticalPageBreaks"), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_VerticalPageBreaks_WithSimplePanel_RowShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 5, 4);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            IXLRange child1 = ws.Range(2, 3, 5, 3);
            child1.AddToNamed("ChildRange1", XLScope.Worksheet);

            IXLRange child2 = ws.Range(2, 4, 5, 4);
            child2.AddToNamed("ChildRange2", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(3, 2).Value = "{di:Date}";

            ws.Cell(3, 3).Value = "{di:Field1}";
            ws.Cell(4, 3).Value = "{di:Field2}";
            ws.Cell(5, 3).Value = "{di:parent:Sum}";

            ws.Cell(2, 4).Value = "{VertPageBreak}";

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report, report.TemplateProcessor)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.Row,
            };
            var childPanel1 = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange1"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.Row,
            };
            var childPanel2 = new ExcelNamedPanel(ws.NamedRange("ChildRange2"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
            };
            parentPanel.Children = new[] { childPanel1, childPanel2 };
            parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 5, 12), parentPanel.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_PageBreaks_Test),
                "Test_VerticalPageBreaks"), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void Test_VerticalPageBreaks_WithoutSimplePanel_RowsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 5, 4);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            IXLRange child = ws.Range(2, 3, 5, 3);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(3, 2).Value = "{di:Date}";

            ws.Cell(3, 3).Value = "{di:Field1}";
            ws.Cell(4, 3).Value = "{di:Field2}";
            ws.Cell(5, 3).Value = "{di:parent:Sum}";

            ws.Cell(2, 4).Value = "{VertPageBreak}";

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report, report.TemplateProcessor)
            {
                Type = PanelType.Horizontal,
            };
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.Row,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 5, 12), parentPanel.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_PageBreaks_Test),
                "Test_VerticalPageBreaks"), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}