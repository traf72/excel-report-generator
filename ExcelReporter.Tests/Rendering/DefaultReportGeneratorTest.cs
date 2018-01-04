using System;
using System.Collections.Generic;
using System.Reflection;
using ClosedXML.Excel;
using ExcelReporter.Rendering;
using ExcelReporter.Rendering.Panels.ExcelPanels;
using ExcelReporter.Rendering.TemplateProcessors;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;

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
    }
}