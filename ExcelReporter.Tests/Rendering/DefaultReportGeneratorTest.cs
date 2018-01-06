using ClosedXML.Excel;
using ExcelReporter.Rendering;
using ExcelReporter.Rendering.Panels.ExcelPanels;
using ExcelReporter.Rendering.Providers;
using ExcelReporter.Rendering.TemplateProcessors;
using ExcelReporter.Tests.CustomAsserts;
using ExcelReporter.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelReporter.Tests.Rendering
{
    [TestClass]
    public class DefaultReportGeneratorTest
    {
        [TestMethod]
        public void TestMakePanelsHierarchy()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLRange panel1Range = ws.Range(1, 1, 4, 4);
            IXLRange panel2Range = ws.Range(1, 1, 2, 4);
            panel2Range.AddToNamed("Panel2", XLScope.Worksheet);
            IXLRange panel3Range = ws.Range(2, 1, 2, 4);
            panel3Range.AddToNamed("Panel3", XLScope.Workbook);
            IXLRange panel4Range = ws.Range(5, 1, 6, 5);
            IXLRange panel5Range = ws.Range(6, 1, 6, 5);
            panel5Range.AddToNamed("Panel5", XLScope.Worksheet);
            IXLRange panel6Range = ws.Range(3, 1, 4, 4);
            IXLRange panel7Range = ws.Range(10, 10, 10, 10);
            IXLRange panel8Range = ws.Range(8, 9, 9, 10);
            panel8Range.AddToNamed("Panel8", XLScope.Worksheet);

            var panel1 = new ExcelPanel(panel1Range, new object(), Substitute.For<ITemplateProcessor>());
            var panel2 = new ExcelDataSourcePanel("Stub", ws.NamedRange("Panel2"), new object(), Substitute.For<ITemplateProcessor>());
            var panel3 = new ExcelDataSourcePanel("Stub", wb.NamedRange("Panel3"), new object(), Substitute.For<ITemplateProcessor>());
            var panel4 = new ExcelPanel(panel4Range, new object(), Substitute.For<ITemplateProcessor>());
            var panel5 = new ExcelDataSourceDynamicPanel("Stub", ws.NamedRange("Panel5"), new object(), Substitute.For<ITemplateProcessor>());
            var panel6 = new ExcelPanel(panel6Range, new object(), Substitute.For<ITemplateProcessor>());
            var panel7 = new ExcelPanel(panel7Range, new object(), Substitute.For<ITemplateProcessor>());
            var panel8 = new ExcelTotalsPanel("Stub", ws.NamedRange("Panel8"), new object(), Substitute.For<ITemplateProcessor>());

            IDictionary<string, (IExcelPanel, string)> panelsFlatView = new Dictionary<string, (IExcelPanel, string)>
            {
                ["Panel1"] = (panel1, null),
                ["Panel2"] = (panel2, "Panel1"),
                ["Panel3"] = (panel3, "Panel2"),
                ["Panel4"] = (panel4, null),
                ["Panel5"] = (panel5, "Panel4"),
                ["Panel6"] = (panel6, "Panel1"),
                ["Panel7"] = (panel7, null),
                ["Panel8"] = (panel8, null),
            };

            var reportGenerator = new DefaultReportGenerator(new object());
            MethodInfo method = reportGenerator.GetType().GetMethod("MakePanelsHierarchy", BindingFlags.Instance | BindingFlags.NonPublic);

            var rootPanel = new ExcelPanel(ws.Range(1, 1, 10, 10), new object(), Substitute.For<ITemplateProcessor>());
            method.Invoke(reportGenerator, new object[] { panelsFlatView, rootPanel });

            Assert.AreEqual(4, rootPanel.Children.Count);
            Assert.AreEqual(panel1Range, rootPanel.Children[0].Range);
            Assert.AreEqual(panel4Range, rootPanel.Children[1].Range);
            Assert.AreEqual(panel7Range, rootPanel.Children[2].Range);
            Assert.AreEqual(panel8Range, rootPanel.Children[3].Range);
            Assert.AreEqual(rootPanel, panel1.Parent);
            Assert.AreEqual(rootPanel, panel4.Parent);
            Assert.AreEqual(rootPanel, panel7.Parent);
            Assert.AreEqual(rootPanel, panel8.Parent);
            Assert.IsNull(rootPanel.Parent);

            Assert.AreEqual(2, panel1.Children.Count);
            Assert.AreEqual(panel2Range, panel1.Children[0].Range);
            Assert.AreEqual(panel6Range, panel1.Children[1].Range);
            Assert.AreEqual(panel1, panel2.Parent);
            Assert.AreEqual(panel1, panel6.Parent);

            Assert.AreEqual(1, panel2.Children.Count);
            Assert.AreEqual(panel3Range, panel2.Children[0].Range);
            Assert.AreEqual(panel2, panel3.Parent);

            Assert.AreEqual(1, panel4.Children.Count);
            Assert.AreEqual(panel5Range, panel4.Children[0].Range);
            Assert.AreEqual(panel4, panel5.Parent);
        }

        [TestMethod]
        public void TestMakePanelsHierarchyWithBadParent()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLRange panel1Range = ws.Range(1, 1, 4, 4);
            IXLRange panel2Range = ws.Range(1, 1, 2, 4);
            panel2Range.AddToNamed("Panel2", XLScope.Worksheet);

            var panel1 = new ExcelPanel(panel1Range, new object(), Substitute.For<ITemplateProcessor>());
            var panel2 = new ExcelDataSourcePanel("Stub", ws.NamedRange("Panel2"), new object(), Substitute.For<ITemplateProcessor>());

            IDictionary<string, (IExcelPanel, string)> panelsFlatView = new Dictionary<string, (IExcelPanel, string)>
            {
                ["Panel1"] = (panel1, null),
                ["Panel2"] = (panel2, "panel1"),
            };

            var reportGenerator = new DefaultReportGenerator(new object());
            MethodInfo method = reportGenerator.GetType().GetMethod("MakePanelsHierarchy", BindingFlags.Instance | BindingFlags.NonPublic);

            var rootPanel = new ExcelPanel(ws.Range(1, 1, 10, 10), new object(), Substitute.For<ITemplateProcessor>());
            ExceptionAssert.ThrowsBaseException<InvalidOperationException>(() => method.Invoke(reportGenerator, new object[] { panelsFlatView, rootPanel }),
                "Cannot find parent panel with name \"panel1\" for panel \"Panel2\"");
        }

        [TestMethod]
        public void TestGetPanelsNamedRanges()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLRange panel1Range = ws.Range(1, 1, 4, 4);
            panel1Range.AddToNamed("s_Panel1", XLScope.Worksheet);
            IXLRange panel2Range = ws.Range(1, 1, 2, 4);
            panel2Range.AddToNamed("D_Panel2", XLScope.Worksheet);
            IXLRange panel3Range = ws.Range(2, 1, 2, 4);
            panel3Range.AddToNamed("DYN_Panel3", XLScope.Worksheet);
            IXLRange panel4Range = ws.Range(5, 1, 6, 5);
            panel4Range.AddToNamed("t_Panel4", XLScope.Worksheet);
            IXLRange panel5Range = ws.Range(6, 1, 6, 5);
            panel5Range.AddToNamed("ss_Panel5", XLScope.Worksheet);
            IXLRange panel6Range = ws.Range(3, 1, 4, 4);
            panel6Range.AddToNamed("S_Panel6", XLScope.Worksheet);
            IXLRange panel7Range = ws.Range(10, 10, 10, 10);
            panel7Range.AddToNamed("d-Panel7", XLScope.Worksheet);
            IXLRange panel8Range = ws.Range(8, 9, 9, 10);
            panel8Range.AddToNamed("d_Panel8", XLScope.Worksheet);
            IXLRange panel9Range = ws.Range(11, 11, 11, 11);
            panel9Range.AddToNamed(" d_Panel9 ", XLScope.Worksheet);

            var reportGenerator = new DefaultReportGenerator(new object());
            MethodInfo method = reportGenerator.GetType().GetMethod("GetPanelsNamedRanges", BindingFlags.Instance | BindingFlags.NonPublic);

            var restult = (IList<IXLNamedRange>)method.Invoke(reportGenerator, new object[] { ws.NamedRanges });

            Assert.AreEqual(6, restult.Count);
            Assert.AreEqual("s_Panel1", restult[0].Name);
            Assert.AreEqual("D_Panel2", restult[1].Name);
            Assert.AreEqual("DYN_Panel3", restult[2].Name);
            Assert.AreEqual("t_Panel4", restult[3].Name);
            Assert.AreEqual("S_Panel6", restult[4].Name);
            Assert.AreEqual("d_Panel8", restult[5].Name);
        }

        [TestMethod]
        public void TestRender()
        {
            var report = new TestReport();
            XLWorkbook wb = report.Workbook;
            IXLWorksheet sheet1 = wb.AddWorksheet("Sheet1");
            IXLWorksheet sheet2 = wb.AddWorksheet("Sheet2");
            var reprotGenerator = new TestReportGenerator(report);

            IXLRange parentRange = sheet1.Range(2, 2, 3, 5);
            parentRange.AddToNamed("d_Parent", XLScope.Worksheet);
            IXLNamedRange parentNamedRange = sheet1.NamedRange("d_Parent");
            parentNamedRange.Comment = "DataSource = m:DataProvider:GetIEnumerable()";

            IXLRange childRange = sheet1.Range(3, 2, 3, 5);
            childRange.AddToNamed("d_Child", XLScope.Workbook);
            IXLNamedRange childNamedRange = wb.NamedRange("d_Child");
            childNamedRange.Comment = $"ParentPanel = d_Parent{Environment.NewLine}DataSource = m:DataProvider:GetChildIEnumerable(di:Name)";

            sheet1.Cell(2, 2).Value = "{di:Name}";
            sheet1.Cell(2, 3).Value = "{di:Date}";
            sheet1.Cell(3, 3).Value = "{di:Field1}";
            sheet1.Cell(3, 4).Value = "{di:Field2}";
            sheet1.Cell(3, 5).Value = "{di:parent:Sum}";

            IXLRange simpleRange = sheet1.Range(2, 7, 3, 8);
            simpleRange.AddToNamed("s_Simple", XLScope.Worksheet);

            sheet1.Cell(2, 7).Value = "{p:StrParam}";
            sheet1.Cell(3, 8).Value = "{p:IntParam}";

            IXLRange dynamicRange = sheet2.Range(2, 2, 4, 2);
            dynamicRange.AddToNamed("dyn_Dynamic", XLScope.Workbook);
            IXLNamedRange dynamicNamedRange = wb.NamedRange("dyn_Dynamic");
            dynamicNamedRange.Comment = "DataSource = m:DataProvider:GetAllCustomersDataReader()";

            sheet2.Cell(2, 2).Value = "{Headers}";
            sheet2.Cell(3, 2).Value = "{Data}";
            sheet2.Cell(4, 2).Value = "{Totals}";

            IXLRange totalsRange = sheet2.Range(6, 2, 6, 7);
            totalsRange.AddToNamed("t_Totals", XLScope.Worksheet);
            IXLNamedRange totalsNamedRange = sheet2.NamedRange("t_Totals");
            totalsNamedRange.Comment = "DataSource = m:DataProvider:GetIEnumerable(); BeforeRenderMethodName = TestExcelTotalsPanelBeforeRender; AfterRenderMethodName = TestExcelTotalsPanelAfterRender";

            sheet2.Cell(6, 2).Value = "Plain text";
            sheet2.Cell(6, 3).Value = "{Sum(di:Sum)}";
            sheet2.Cell(6, 4).Value = "{ Custom(DI:Sum, CustomAggregation, PostAggregation)  }";
            sheet2.Cell(6, 5).Value = "{Min(di:Sum)}";
            sheet2.Cell(6, 6).Value = "Text1 {count(Name)} Text2 {avg(di:Sum, , PostAggregationRound)} Text3 {Max(Sum)}";
            sheet2.Cell(6, 7).Value = "{Mix(di:Sum)}";

            sheet2.Cell(10, 10).Value = "Plain text";
            sheet2.Cell(1, 1).Value = " { m:Format ( p:DateParam ) } ";
            sheet2.Cell(7, 1).Value = "{P:BoolParam}";

            reprotGenerator.Render(wb);

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DefaultReportGeneratorTest), nameof(TestRender)), wb);

            //wb.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestRenderPartialWorksheets()
        {
            var report = new TestReport();
            XLWorkbook wb = report.Workbook;
            IXLWorksheet sheet1 = wb.AddWorksheet("Sheet1");
            IXLWorksheet sheet2 = wb.AddWorksheet("Sheet2");
            var reprotGenerator = new TestReportGenerator(report);

            IXLRange parentRange = sheet1.Range(2, 2, 3, 5);
            parentRange.AddToNamed("d_Parent", XLScope.Worksheet);
            IXLNamedRange parentNamedRange = sheet1.NamedRange("d_Parent");
            parentNamedRange.Comment = "DataSource = m:DataProvider:GetIEnumerable()";

            IXLRange childRange = sheet1.Range(3, 2, 3, 5);
            childRange.AddToNamed("d_Child", XLScope.Workbook);
            IXLNamedRange childNamedRange = wb.NamedRange("d_Child");
            childNamedRange.Comment = $"ParentPanel = d_Parent{Environment.NewLine}DataSource = m:DataProvider:GetChildIEnumerable(di:Name)";

            sheet1.Cell(2, 2).Value = "{di:Name}";
            sheet1.Cell(2, 3).Value = "{di:Date}";
            sheet1.Cell(3, 3).Value = "{di:Field1}";
            sheet1.Cell(3, 4).Value = "{di:Field2}";
            sheet1.Cell(3, 5).Value = "{di:parent:Sum}";

            IXLRange simpleRange = sheet1.Range(2, 7, 3, 8);
            simpleRange.AddToNamed("s_Simple", XLScope.Worksheet);

            sheet1.Cell(2, 7).Value = "{p:StrParam}";
            sheet1.Cell(3, 8).Value = "{p:IntParam}";

            IXLRange dynamicRange = sheet2.Range(2, 2, 4, 2);
            dynamicRange.AddToNamed("dyn_Dynamic", XLScope.Workbook);
            IXLNamedRange dynamicNamedRange = wb.NamedRange("dyn_Dynamic");
            dynamicNamedRange.Comment = "DataSource = m:DataProvider:GetAllCustomersDataReader()";

            sheet2.Cell(2, 2).Value = "{Headers}";
            sheet2.Cell(3, 2).Value = "{Data}";
            sheet2.Cell(4, 2).Value = "{Totals}";

            IXLRange totalsRange = sheet2.Range(6, 2, 6, 7);
            totalsRange.AddToNamed("t_Totals", XLScope.Worksheet);
            IXLNamedRange totalsNamedRange = sheet2.NamedRange("t_Totals");
            totalsNamedRange.Comment = "DataSource = m:DataProvider:GetIEnumerable(); BeforeRenderMethodName = TestExcelTotalsPanelBeforeRender; AfterRenderMethodName = TestExcelTotalsPanelAfterRender";

            sheet2.Cell(6, 2).Value = "Plain text";
            sheet2.Cell(6, 3).Value = "{Sum(di:Sum)}";
            sheet2.Cell(6, 4).Value = "{ Custom(DI:Sum, CustomAggregation, PostAggregation)  }";
            sheet2.Cell(6, 5).Value = "{Min(di:Sum)}";
            sheet2.Cell(6, 6).Value = "Text1 {count(Name)} Text2 {avg(di:Sum, , PostAggregationRound)} Text3 {Max(Sum)}";
            sheet2.Cell(6, 7).Value = "{Mix(di:Sum)}";

            sheet2.Cell(10, 10).Value = "Plain text";
            sheet2.Cell(1, 1).Value = " { m:Format ( p:DateParam ) } ";
            sheet2.Cell(7, 1).Value = "{P:BoolParam}";

            reprotGenerator.Render(wb, new[] { sheet1 });

            ExcelAssert.AreWorkbooksContentEquals(TestHelper.GetExpectedWorkbook(nameof(DefaultReportGeneratorTest), nameof(TestRenderPartialWorksheets)), wb);

            //wb.SaveAs("test.xlsx");
        }

        private class TestReportGenerator : DefaultReportGenerator
        {
            private ITypeProvider _typeProvider;

            public TestReportGenerator(object report) : base(report)
            {
            }

            public override ITypeProvider TypeProvider => _typeProvider ?? (_typeProvider = new DefaultTypeProvider(new[] { Assembly.GetExecutingAssembly() }, _report.GetType()));
        }
    }
}