using ClosedXML.Excel;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;
using NUnit.Framework;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests.DataSourcePanelRenderTests
{
    
    public class DataSourcePanelRender_WithGrouping_MixedPanels_Test
    {
        [Test]
        public void TestMultipleVerticalPanelsGrouping()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 8, 7);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange simpleRange1 = ws.Range(3, 3, 3, 4);
            simpleRange1.AddToNamed("simpleRange1");

            simpleRange1.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            simpleRange1.Style.Border.OutsideBorderColor = XLColor.Brown;

            IXLRange child = ws.Range(4, 2, 7, 7);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            child.Range(1, 2, 4, 6).Style.Border.TopBorder = XLBorderStyleValues.Thin;
            child.Range(1, 2, 4, 6).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            child.Range(1, 2, 4, 6).Style.Border.TopBorderColor = XLColor.Red;
            child.Range(1, 2, 4, 6).Style.Border.LeftBorderColor = XLColor.Red;
            child.Range(1, 2, 4, 6).Style.Border.InsideBorder = XLBorderStyleValues.None;

            IXLRange childOfChild = ws.Range(5, 2, 6, 7);
            childOfChild.AddToNamed("ChildOfChildRange");

            childOfChild.Range(1, 3, 2, 6).Style.Border.TopBorder = XLBorderStyleValues.Thin;
            childOfChild.Range(1, 3, 2, 6).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            childOfChild.Range(1, 3, 2, 6).Style.Border.TopBorderColor = XLColor.Green;
            childOfChild.Range(1, 3, 2, 6).Style.Border.LeftBorderColor = XLColor.Green;
            childOfChild.Range(1, 3, 2, 6).Style.Border.InsideBorder = XLBorderStyleValues.None;

            IXLRange simpleRange2 = ws.Range(6, 4, 6, 7);
            simpleRange2.AddToNamed("simpleRange2", XLScope.Worksheet);

            simpleRange2.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            simpleRange2.Style.Border.TopBorderColor = XLColor.Orange;

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";

            ws.Cell(3, 3).Value = "{p:StrParam}";
            ws.Cell(3, 4).Value = "{di:Sum}";
            ws.Cell(3, 5).Value = "{p:IntParam}";

            ws.Cell(4, 3).Value = "{di:Field1}";
            ws.Cell(4, 4).Value = "{di:Field2}";
            ws.Cell(4, 5).Value = "{di:parent:Sum}";
            ws.Cell(4, 6).Value = "{di:parent:Contacts}";

            ws.Cell(5, 4).Value = "{di:Field1}";
            ws.Cell(5, 5).Value = "{di:Field2}";
            ws.Cell(5, 6).Value = "{di:parent:Field1}";
            ws.Cell(5, 7).Value = "{di:parent:parent:Contacts.Phone}";

            ws.Cell(6, 5).Value = "{p:DateParam}";
            ws.Cell(6, 6).Value = "{di:parent:Field2}";
            ws.Cell(6, 7).Value = "{di:parent:parent:Contacts.Fax}";

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report, report.TemplateProcessor)
            {
                BeforeDataItemRenderMethodName = "BeforeRenderParentDataSourcePanel",
                AfterDataItemRenderMethodName = "AfterRenderParentDataSourcePanelChildBottom",
            };
            var simplePanel1 = new ExcelNamedPanel(ws.Workbook.NamedRange("simpleRange1"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
            };
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
                AfterDataItemRenderMethodName = "AfterRenderChildDataSourcePanel",
            };
            var childOfChildPanel = new ExcelDataSourcePanel("di:Children", ws.Workbook.NamedRange("ChildOfChildRange"), report, report.TemplateProcessor)
            {
                Parent = childPanel
            };
            var simplePanel2 = new ExcelNamedPanel(ws.NamedRange("simpleRange2"), report, report.TemplateProcessor)
            {
                Parent = childOfChildPanel,
            };

            childOfChildPanel.Children = new[] { simplePanel2 };
            childPanel.Children = new[] { childOfChildPanel };
            parentPanel.Children = new[] { childPanel, simplePanel1 };
            parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 20, 7), parentPanel.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_MixedPanels_Test),
                nameof(TestMultipleVerticalPanelsGrouping)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [Test]
        public void TestMultipleHorizontalPanelsGrouping()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 7, 6);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange simpleRange1 = ws.Range(3, 3, 4, 3);
            simpleRange1.AddToNamed("simpleRange1");

            IXLRange child = ws.Range(2, 4, 7, 6);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            IXLRange childOfChild = ws.Range(2, 5, 7, 6);
            childOfChild.AddToNamed("ChildOfChildRange");

            IXLRange simpleRange2 = ws.Range(5, 6, 7, 6);
            simpleRange2.AddToNamed("simpleRange2", XLScope.Worksheet);

            simpleRange2.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            simpleRange2.Style.Border.OutsideBorderColor = XLColor.Orange;

            childOfChild.Range(3, 1, 6, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            childOfChild.Range(3, 1, 6, 2).Style.Border.OutsideBorderColor = XLColor.Green;

            child.Range(2, 1, 6, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            child.Range(2, 1, 6, 3).Style.Border.OutsideBorderColor = XLColor.Red;

            simpleRange1.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            simpleRange1.Style.Border.OutsideBorderColor = XLColor.Brown;

            parentRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            parentRange.Style.Border.OutsideBorderColor = XLColor.Black;

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(3, 2).Value = "{di:Date}";

            ws.Cell(3, 3).Value = "{p:StrParam}";
            ws.Cell(4, 3).Value = "{di:Sum}";
            ws.Cell(5, 3).Value = "{p:IntParam}";

            ws.Cell(3, 4).Value = "{di:Field1}";
            ws.Cell(4, 4).Value = "{di:Field2}";
            ws.Cell(5, 4).Value = "{di:parent:Sum}";
            ws.Cell(6, 4).Value = "{di:parent:Contacts}";

            ws.Cell(4, 5).Value = "{di:Field1}";
            ws.Cell(5, 5).Value = "{di:Field2}";
            ws.Cell(6, 5).Value = "{di:parent:Field1}";
            ws.Cell(7, 5).Value = "{di:parent:parent:Contacts.Phone}";

            ws.Cell(5, 6).Value = "{p:DateParam}";
            ws.Cell(6, 6).Value = "{di:parent:Field2}";
            ws.Cell(7, 6).Value = "{di:parent:parent:Contacts.Fax}";

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report, report.TemplateProcessor)
            {
                Type = PanelType.Horizontal,
            };
            var simplePanel1 = new ExcelNamedPanel(ws.Workbook.NamedRange("simpleRange1"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
            };
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
            };
            var childOfChildPanel = new ExcelDataSourcePanel("di:Children", ws.Workbook.NamedRange("ChildOfChildRange"), report, report.TemplateProcessor)
            {
                Parent = childPanel,
                Type = PanelType.Horizontal,
            };
            var simplePanel2 = new ExcelNamedPanel(ws.NamedRange("simpleRange2"), report, report.TemplateProcessor)
            {
                Parent = childOfChildPanel,
            };

            childOfChildPanel.Children = new[] { simplePanel2 };
            childPanel.Children = new[] { childOfChildPanel };
            parentPanel.Children = new[] { childPanel, simplePanel1 };
            parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 7, 20), parentPanel.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_MixedPanels_Test),
                nameof(TestMultipleHorizontalPanelsGrouping)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [Test]
        public void TestHorizontalInVerticalPanelsGrouping()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 4, 3);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange child = ws.Range(2, 3, 4, 3);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";

            ws.Cell(3, 3).Value = "{di:Field1}";
            ws.Cell(4, 3).Value = "{di:Field2}";

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report, report.TemplateProcessor);
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
                Type = PanelType.Horizontal,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 10, 5), parentPanel.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_MixedPanels_Test),
                nameof(TestHorizontalInVerticalPanelsGrouping)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [Test]
        public void TestVerticalInHorizontalPanelsGrouping()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange parentRange = ws.Range(2, 2, 3, 4);
            parentRange.AddToNamed("ParentRange", XLScope.Worksheet);

            IXLRange child = ws.Range(3, 3, 3, 4);
            child.AddToNamed("ChildRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";

            ws.Cell(3, 3).Value = "{di:Field1}";
            ws.Cell(3, 4).Value = "{di:Field2}";

            var parentPanel = new ExcelDataSourcePanel("m:DataProvider:GetIEnumerable()", ws.NamedRange("ParentRange"), report, report.TemplateProcessor)
            {
                Type = PanelType.Horizontal,
            };
            var childPanel = new ExcelDataSourcePanel("m:DataProvider:GetChildIEnumerable(di:Name)", ws.NamedRange("ChildRange"), report, report.TemplateProcessor)
            {
                Parent = parentPanel,
            };
            parentPanel.Children = new[] { childPanel };
            parentPanel.Render();

            Assert.AreEqual(ws.Range(2, 2, 5, 10), parentPanel.ResultRange);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DataSourcePanelRender_WithGrouping_MixedPanels_Test),
                nameof(TestVerticalInHorizontalPanelsGrouping)), ws.Workbook);

            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}