using ClosedXML.Excel;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests;
using NUnit.Framework;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels
{
    
    public class PanelsExpandTest
    {
        /// <summary>
        /// The child panel occupies the entire width of the parent (cells are moved)
        /// </summary>
        [Test]
        public void TestExpandSimplePanel_ChildAcrossWidth_ChildCenter_CellsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 4, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            IXLRange childRange = ws.Range(3, 2, 3, 5);
            childRange.AddToNamed("ChildRange", XLScope.Worksheet);

            childRange.Style.Border.OutsideBorder = XLBorderStyleValues.Dashed;
            childRange.Style.Border.OutsideBorderColor = XLColor.Blue;

            ws.Cell(2, 2).Value = "{p:StrParam}";

            ws.Cell(3, 3).Value = "{di:Name}";
            ws.Cell(3, 4).Value = "{di:Date}";
            ws.Cell(3, 5).Value = "{di:Sum}";

            var parentPanel = new ExcelPanel(parentRange, report, report.TemplateProcessor);
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 6, 5), parentPanel.ResultRange);

            //report.Workbook.SaveAs("test.xlsx");
        }

        /// <summary>
        /// The child panel occupies the entire width of the parent (the whole row is moved)
        /// </summary>
        [Test]
        public void TestExpandSimplePanel_ChildAcrossWidth_ChildCenter_RowShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 4, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            IXLRange childRange = ws.Range(3, 2, 3, 5);
            childRange.AddToNamed("ChildRange", XLScope.Worksheet);

            childRange.Style.Border.OutsideBorder = XLBorderStyleValues.Dashed;
            childRange.Style.Border.OutsideBorderColor = XLColor.Blue;

            ws.Cell(2, 2).Value = "{p:StrParam}";

            ws.Cell(3, 3).Value = "{di:Name}";
            ws.Cell(3, 4).Value = "{di:Date}";
            ws.Cell(3, 5).Value = "{di:Sum}";

            var parentPanel = new ExcelPanel(parentRange, report, report.TemplateProcessor);
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
                ShiftType = ShiftType.Row,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 6, 5), parentPanel.ResultRange);

            //report.Workbook.SaveAs("test.xlsx");
        }

        /// <summary>
        /// The child panel occupies the entire width of the parent (without any shift)
        /// </summary>
        [Test]
        public void TestExpandSimplePanel_ChildAcrossWidth_ChildCenter_NoShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 4, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            IXLRange childRange = ws.Range(3, 2, 3, 5);
            childRange.AddToNamed("ChildRange", XLScope.Worksheet);

            childRange.Style.Border.OutsideBorder = XLBorderStyleValues.Dashed;
            childRange.Style.Border.OutsideBorderColor = XLColor.Blue;

            ws.Cell(2, 2).Value = "{p:StrParam}";

            ws.Cell(3, 3).Value = "{di:Name}";
            ws.Cell(3, 4).Value = "{di:Date}";
            ws.Cell(3, 5).Value = "{di:Sum}";

            var parentPanel = new ExcelPanel(parentRange, report, report.TemplateProcessor);
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
                ShiftType = ShiftType.NoShift,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 5, 5), parentPanel.ResultRange);

            //report.Workbook.SaveAs("test.xlsx");
        }

        /// <summary>
        /// The child panel occupies the entire width of the parent (cells are moved)
        /// </summary>
        [Test]
        public void TestExpandSimplePanel_ChildAcrossWidth_ChildBottom_CellsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 3, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            IXLRange childRange = ws.Range(3, 2, 3, 5);
            childRange.AddToNamed("ChildRange", XLScope.Worksheet);

            childRange.Style.Border.OutsideBorder = XLBorderStyleValues.Dashed;
            childRange.Style.Border.OutsideBorderColor = XLColor.Blue;

            ws.Cell(2, 2).Value = "{p:StrParam}";

            ws.Cell(3, 3).Value = "{di:Name}";
            ws.Cell(3, 4).Value = "{di:Date}";
            ws.Cell(3, 5).Value = "{di:Sum}";

            var parentPanel = new ExcelPanel(parentRange, report, report.TemplateProcessor);
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 5, 5), parentPanel.ResultRange);

            //report.Workbook.SaveAs("test.xlsx");
        }

        /// <summary>
        /// The child panel occupies the entire width of the parent (the whole row is moved)
        /// </summary>
        [Test]
        public void TestExpandSimplePanel_ChildAcrossWidth_ChildBottom_RowShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 3, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            IXLRange childRange = ws.Range(3, 2, 3, 5);
            childRange.AddToNamed("ChildRange", XLScope.Worksheet);

            childRange.Style.Border.OutsideBorder = XLBorderStyleValues.Dashed;
            childRange.Style.Border.OutsideBorderColor = XLColor.Blue;

            ws.Cell(2, 2).Value = "{p:StrParam}";

            ws.Cell(3, 3).Value = "{di:Name}";
            ws.Cell(3, 4).Value = "{di:Date}";
            ws.Cell(3, 5).Value = "{di:Sum}";

            var parentPanel = new ExcelPanel(parentRange, report, report.TemplateProcessor);
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
                ShiftType = ShiftType.Row,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 5, 5), parentPanel.ResultRange);

            //report.Workbook.SaveAs("test.xlsx");
        }

        /// <summary>
        /// The child panel occupies the entire width of the parent (without any shift)
        /// </summary>
        [Test]
        public void TestExpandSimplePanel_ChildAcrossWidth_ChildBottom_NoShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 3, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            IXLRange childRange = ws.Range(3, 2, 3, 5);
            childRange.AddToNamed("ChildRange", XLScope.Worksheet);

            childRange.Style.Border.OutsideBorder = XLBorderStyleValues.Dashed;
            childRange.Style.Border.OutsideBorderColor = XLColor.Blue;

            ws.Cell(2, 2).Value = "{p:StrParam}";

            ws.Cell(3, 3).Value = "{di:Name}";
            ws.Cell(3, 4).Value = "{di:Date}";
            ws.Cell(3, 5).Value = "{di:Sum}";

            var parentPanel = new ExcelPanel(parentRange, report, report.TemplateProcessor);
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
                ShiftType = ShiftType.NoShift,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 5, 5), parentPanel.ResultRange);

            //report.Workbook.SaveAs("test.xlsx");
        }

        /// <summary>
        /// he child panel occupies the entire width of the parent (cells are moved)
        /// </summary>
        [Test]
        public void TestExpandSimplePanel_ChildNotAcrossWidth_CellsShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 4, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            IXLRange childRange = ws.Range(3, 3, 3, 5);
            childRange.AddToNamed("ChildRange", XLScope.Worksheet);

            childRange.Style.Border.OutsideBorder = XLBorderStyleValues.Dashed;
            childRange.Style.Border.OutsideBorderColor = XLColor.Blue;

            ws.Cell(2, 2).Value = "{p:StrParam}";

            ws.Cell(3, 3).Value = "{di:Name}";
            ws.Cell(3, 4).Value = "{di:Date}";
            ws.Cell(3, 5).Value = "{di:Sum}";

            var parentPanel = new ExcelPanel(parentRange, report, report.TemplateProcessor);
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 5, 5), parentPanel.ResultRange);

            //report.Workbook.SaveAs("test.xlsx");
        }

        /// <summary>
        /// The child panel occupies the entire width of the parent (the whole row is moved)
        /// </summary>
        [Test]
        public void TestExpandSimplePanel_ChildNotAcrossWidth_RowShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 4, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            IXLRange childRange = ws.Range(3, 3, 3, 5);
            childRange.AddToNamed("ChildRange", XLScope.Worksheet);

            childRange.Style.Border.OutsideBorder = XLBorderStyleValues.Dashed;
            childRange.Style.Border.OutsideBorderColor = XLColor.Blue;

            ws.Cell(2, 2).Value = "{p:StrParam}";

            ws.Cell(3, 3).Value = "{di:Name}";
            ws.Cell(3, 4).Value = "{di:Date}";
            ws.Cell(3, 5).Value = "{di:Sum}";

            var parentPanel = new ExcelPanel(parentRange, report, report.TemplateProcessor);
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
                ShiftType =  ShiftType.Row,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 6, 5), parentPanel.ResultRange);

            //report.Workbook.SaveAs("test.xlsx");
        }

        /// <summary>
        /// The child panel occupies the entire width of the parent (without any shift)
        /// </summary>
        [Test]
        public void TestExpandSimplePanel_ChildNotAcrossWidth_NoShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 4, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            IXLRange childRange = ws.Range(3, 3, 3, 5);
            childRange.AddToNamed("ChildRange", XLScope.Worksheet);

            childRange.Style.Border.OutsideBorder = XLBorderStyleValues.Dashed;
            childRange.Style.Border.OutsideBorderColor = XLColor.Blue;

            ws.Cell(2, 2).Value = "{p:StrParam}";

            ws.Cell(3, 3).Value = "{di:Name}";
            ws.Cell(3, 4).Value = "{di:Date}";
            ws.Cell(3, 5).Value = "{di:Sum}";

            var parentPanel = new ExcelPanel(parentRange, report, report.TemplateProcessor);
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
                ShiftType = ShiftType.NoShift,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 5, 5), parentPanel.ResultRange);

            //report.Workbook.SaveAs("test.xlsx");
        }

        /// <summary>
        /// The child panel occupies the entire width of the parent (the whole row is moved)
        /// </summary>
        [Test]
        public void TestExpandDataPanel_ChildNotAcrossWidth_ChildCenter_RowShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 4, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            IXLRange childRange = ws.Range(3, 3, 3, 5);
            childRange.AddToNamed("ChildRange", XLScope.Worksheet);

            childRange.Style.Border.OutsideBorder = XLBorderStyleValues.Dashed;
            childRange.Style.Border.OutsideBorderColor = XLColor.Blue;

            ws.Cell(2, 2).Value = "{di:Name}";

            ws.Cell(3, 3).Value = "{di:Field1}";
            ws.Cell(3, 4).Value = "{di:Field2}";

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report, report.TemplateProcessor);
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
                ShiftType = ShiftType.Row,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 12, 5), parentPanel.ResultRange);

            //report.Workbook.SaveAs("test.xlsx");
        }

        /// <summary>
        /// The child panel occupies the entire width of the parent (the whole row is moved)
        /// </summary>
        [Test]
        public void TestExpandDataPanel_ChildNotAcrossWidth_ChildBottom_RowShift()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 3, 5);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            IXLRange childRange = ws.Range(3, 3, 3, 5);
            childRange.AddToNamed("ChildRange", XLScope.Worksheet);

            childRange.Style.Border.OutsideBorder = XLBorderStyleValues.Dashed;
            childRange.Style.Border.OutsideBorderColor = XLColor.Blue;

            ws.Cell(2, 2).Value = "{di:Name}";

            ws.Cell(3, 3).Value = "{di:Field1}";
            ws.Cell(3, 4).Value = "{di:Field2}";

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report, report.TemplateProcessor);
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
                ShiftType = ShiftType.Row,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 9, 5), parentPanel.ResultRange);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}