using ClosedXML.Excel;
using ExcelReporter.Rendering.Panels.ExcelPanels;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace ExcelReporter.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests
{
    [TestClass]
    public class PanelRenderTest
    {
        [TestMethod]
        public void TestPanelRender()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(1, 1, 4, 5);

            ws.Cell(1, 1).Value = "{p:StrParam}";
            ws.Cell(1, 2).Value = "{p:IntParam}";
            ws.Cell(1, 3).Value = "{p:DateParam}";
            ws.Cell(1, 4).Value = "{P:BoolParam}";
            ws.Cell(1, 5).Value = "{p:TimeSpanParam}";
            ws.Cell(2, 1).Value = " { p:StrParam } ";
            ws.Cell(2, 2).Value = "Plain text";
            ws.Cell(2, 3).Value = "{Plain text}";
            ws.Cell(2, 4).Value = " { m:Format ( p:DateParam ) } ";
            ws.Cell(2, 5).Value = "''{m:Format(p:DateParam)}";
            ws.Cell(3, 1).Value = "Int: { p:IntParam }. Str: {p:ComplexTypeParam.StrParam}. FormattedDate: {M:Format(p:DateParam)}";
            ws.Cell(3, 2).Value = "''{m:Format(m:DateTime:AddDays(p:ComplexTypeParam.IntParam), \"yyyy-MM-dd\")}";
            ws.Cell(3, 3).Value = "''{m:Format(m:AddDays(p:DateParam, 5), ddMMyyyy)}";
            ws.Cell(3, 4).Value = "''{m:Format(m:AddDays(p:DateParam, -2), dd.MM.yyyy)}";
            ws.Cell(3, 5).Value = "''{m:Format(m:AddDays(p:DateParam, [int]-3), \"dd.MM.yyyy HH:mm:ss\")}";
            ws.Cell(4, 1).Value = "{m:TestReport:Counter()}";
            ws.Cell(4, 2).Value = "{ m:TestReport : Counter ( ) }";
            ws.Cell(4, 3).Value = "{m:Counter()}";
            ws.Cell(5, 1).Value = "{p:StrParam}";
            ws.Cell(5, 2).Value = "{m:Counter()}";
            ws.Cell(6, 1).Value = "Plain text outside range";

            var panel = new ExcelPanel(range, report);
            panel.Render();

            Assert.AreEqual(21, ws.CellsUsed().Count());
            Assert.AreEqual("String parameter", ws.Cell(1, 1).Value);
            Assert.AreEqual(10d, ws.Cell(1, 2).Value);
            Assert.AreEqual(new DateTime(2017, 10, 25), ws.Cell(1, 3).Value);
            Assert.AreEqual(true, ws.Cell(1, 4).Value);
            Assert.AreEqual(new TimeSpan(36500, 22, 30, 40), ws.Cell(1, 5).Value);
            Assert.AreEqual(" String parameter ", ws.Cell(2, 1).Value);
            Assert.AreEqual("Plain text", ws.Cell(2, 2).Value);
            Assert.AreEqual("{Plain text}", ws.Cell(2, 3).Value);
            Assert.AreEqual(20171025d, ws.Cell(2, 4).Value);
            Assert.AreEqual("20171025", ws.Cell(2, 5).Value);
            Assert.AreEqual("Int: 10. Str: Complex type string parameter. FormattedDate: 20171025", ws.Cell(3, 1).Value);
            Assert.AreEqual("0001-01-12", ws.Cell(3, 2).Value);
            Assert.AreEqual("30102017", ws.Cell(3, 3).Value);
            Assert.AreEqual("23.10.2017", ws.Cell(3, 4).Value);
            Assert.AreEqual("22.10.2017 00:00:00", ws.Cell(3, 5).Value);

            Assert.AreEqual(1d, ws.Cell(4, 1).Value);
            Assert.AreEqual(2d, ws.Cell(4, 2).Value);
            Assert.AreEqual(3d, ws.Cell(4, 3).Value);
            Assert.IsTrue(ws.Cell(4, 4).IsEmpty());
            Assert.IsTrue(ws.Cell(4, 5).IsEmpty());

            Assert.AreEqual("{p:StrParam}", ws.Cell(5, 1).Value);
            Assert.AreEqual("{m:Counter()}", ws.Cell(5, 2).Value);
            Assert.AreEqual("Plain text outside range", ws.Cell(6, 1).Value);

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestCancelPanelRender()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(1, 1, 1, 2);

            ws.Cell(1, 1).Value = "{p:StrParam}";
            ws.Cell(1, 2).Value = "{p:IntParam}";

            var panel = new ExcelPanel(range, report) { BeforeRenderMethodName = "CancelPanelRender" };
            panel.Render();

            Assert.AreEqual(2, ws.CellsUsed().Count());
            Assert.AreEqual("{p:StrParam}", ws.Cell(1, 1).Value);
            Assert.AreEqual("{p:IntParam}", ws.Cell(1, 2).Value);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestPanelRenderEvents()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(1, 1, 1, 2);

            ws.Cell(1, 1).Value = "{p:StrParam}";
            ws.Cell(1, 2).Value = "{p:IntParam}";

            var panel = new ExcelPanel(range, report)
            {
                BeforeRenderMethodName = "TestExcelPanelBeforeRender",
                AfterRenderMethodName = "TestExcelPanelAfterRender",
            };
            panel.Render();

            Assert.AreEqual(2, ws.CellsUsed().Count());
            Assert.IsTrue((bool)ws.Cell(1, 1).Value);
            Assert.AreEqual(11d, ws.Cell(1, 2).Value);

            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestNamedPanelRender()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(1, 1, 4, 5);
            range.AddToNamed("NamedPanel");

            ws.Cell(1, 1).Value = "{p:StrParam}";
            ws.Cell(1, 2).Value = "{p:IntParam}";
            ws.Cell(1, 3).Value = "{p:DateParam}";
            ws.Cell(1, 4).Value = "{p:BoolParam}";
            ws.Cell(1, 5).Value = "{p:TimeSpanParam}";
            ws.Cell(2, 1).Value = "{ p:StrParam }";
            ws.Cell(2, 2).Value = "Plain text";
            ws.Cell(2, 3).Value = "{Plain text}";
            ws.Cell(2, 4).Value = " { m:Format ( p:DateParam ) } ";
            ws.Cell(2, 5).Value = "''{m:Format(p:DateParam)}";
            ws.Cell(3, 1).Value = "Int: { p:IntParam }. Str: {p:ComplexTypeParam.StrParam}. FormattedDate: {m:Format(p:DateParam)}";
            ws.Cell(3, 2).Value = "''{m:Format(m:DateTime:AddDays(p:ComplexTypeParam.IntParam), \"yyyy-MM-dd\")}";
            ws.Cell(3, 3).Value = "''{m:Format(m:AddDays(p:DateParam, 5), ddMMyyyy)}";
            ws.Cell(3, 4).Value = "''{m:Format(m:AddDays(p:DateParam, -2), dd.MM.yyyy)}";
            ws.Cell(3, 5).Value = "''{m:Format(m:AddDays(p:DateParam, [int]-3), \"dd.MM.yyyy HH:mm:ss\")}";
            ws.Cell(4, 1).Value = "{m:TestReport:Counter()}";
            ws.Cell(4, 2).Value = "{ m:TestReport : Counter ( ) }";
            ws.Cell(4, 3).Value = "{m:Counter()}";
            ws.Cell(5, 1).Value = "{p:StrParam}";
            ws.Cell(5, 2).Value = "{m:Counter()}";
            ws.Cell(6, 1).Value = "Plain text outside range";

            var panel = new ExcelNamedPanel(ws.Workbook.NamedRange("NamedPanel"), report);
            panel.Render();

            Assert.AreEqual(21, ws.CellsUsed().Count());
            Assert.AreEqual("String parameter", ws.Cell(1, 1).Value);
            Assert.AreEqual(10d, ws.Cell(1, 2).Value);
            Assert.AreEqual(new DateTime(2017, 10, 25), ws.Cell(1, 3).Value);
            Assert.AreEqual(true, ws.Cell(1, 4).Value);
            Assert.AreEqual(new TimeSpan(36500, 22, 30, 40), ws.Cell(1, 5).Value);
            Assert.AreEqual("String parameter", ws.Cell(2, 1).Value);
            Assert.AreEqual("Plain text", ws.Cell(2, 2).Value);
            Assert.AreEqual("{Plain text}", ws.Cell(2, 3).Value);
            Assert.AreEqual(20171025d, ws.Cell(2, 4).Value);
            Assert.AreEqual("20171025", ws.Cell(2, 5).Value);
            Assert.AreEqual("Int: 10. Str: Complex type string parameter. FormattedDate: 20171025", ws.Cell(3, 1).Value);
            Assert.AreEqual("0001-01-12", ws.Cell(3, 2).Value);
            Assert.AreEqual("30102017", ws.Cell(3, 3).Value);
            Assert.AreEqual("23.10.2017", ws.Cell(3, 4).Value);
            Assert.AreEqual("22.10.2017 00:00:00", ws.Cell(3, 5).Value);

            Assert.AreEqual(1d, ws.Cell(4, 1).Value);
            Assert.AreEqual(2d, ws.Cell(4, 2).Value);
            Assert.AreEqual(3d, ws.Cell(4, 3).Value);
            Assert.IsTrue(ws.Cell(4, 4).IsEmpty());
            Assert.IsTrue(ws.Cell(4, 5).IsEmpty());

            Assert.AreEqual("{p:StrParam}", ws.Cell(5, 1).Value);
            Assert.AreEqual("{m:Counter()}", ws.Cell(5, 2).Value);
            Assert.AreEqual("Plain text outside range", ws.Cell(6, 1).Value);

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());
            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            IXLRange range2 = ws.Range(8, 1, 11, 5);
            range2.AddToNamed("NamedPanel2", XLScope.Worksheet);

            ws.Cell(8, 1).Value = "{p:StrParam}";
            ws.Cell(8, 2).Value = "{p:IntParam}";
            ws.Cell(8, 3).Value = "{p:DateParam}";
            ws.Cell(8, 4).Value = "{p:BoolParam}";
            ws.Cell(8, 5).Value = "{p:TimeSpanParam}";
            ws.Cell(9, 1).Value = "{ p:StrParam }";
            ws.Cell(9, 2).Value = "Plain text";
            ws.Cell(9, 3).Value = "{Plain text}";
            ws.Cell(9, 4).Value = " { m:Format ( p:DateParam ) } ";
            ws.Cell(9, 5).Value = "''{m:Format(p:DateParam)}";
            ws.Cell(10, 1).Value = "Int: { p:IntParam }. Str: {p:ComplexTypeParam.StrParam}. FormattedDate: {m:Format(p:DateParam)}";
            ws.Cell(10, 2).Value = "''{m:Format(m:DateTime:AddDays(p:ComplexTypeParam.IntParam), \"yyyy-MM-dd\")}";
            ws.Cell(10, 3).Value = "''{m:Format(m:AddDays(p:DateParam, 5), ddMMyyyy)}";
            ws.Cell(10, 4).Value = "''{m:Format(m:AddDays(p:DateParam, -2), dd.MM.yyyy)}";
            ws.Cell(10, 5).Value = "''{m:Format(m:AddDays(p:DateParam, [int]-3), \"dd.MM.yyyy HH:mm:ss\")}";
            ws.Cell(11, 1).Value = "{m:TestReport:Counter()}";
            ws.Cell(11, 2).Value = "{ m:TestReport : Counter ( ) }";
            ws.Cell(11, 3).Value = "{m:Counter()}";
            ws.Cell(12, 1).Value = "{p:StrParam}";
            ws.Cell(12, 2).Value = "{m:Counter()}";
            ws.Cell(13, 1).Value = "Plain text outside range";

            panel = new ExcelNamedPanel(ws.NamedRange("NamedPanel2"), report);
            panel.Render();

            Assert.AreEqual(42, ws.CellsUsed().Count());
            Assert.AreEqual("String parameter", ws.Cell(8, 1).Value);
            Assert.AreEqual(10d, ws.Cell(8, 2).Value);
            Assert.AreEqual(new DateTime(2017, 10, 25), ws.Cell(8, 3).Value);
            Assert.AreEqual(true, ws.Cell(8, 4).Value);
            Assert.AreEqual(new TimeSpan(36500, 22, 30, 40), ws.Cell(8, 5).Value);
            Assert.AreEqual("String parameter", ws.Cell(9, 1).Value);
            Assert.AreEqual("Plain text", ws.Cell(9, 2).Value);
            Assert.AreEqual("{Plain text}", ws.Cell(9, 3).Value);
            Assert.AreEqual(20171025d, ws.Cell(9, 4).Value);
            Assert.AreEqual("20171025", ws.Cell(9, 5).Value);
            Assert.AreEqual("Int: 10. Str: Complex type string parameter. FormattedDate: 20171025", ws.Cell(10, 1).Value);
            Assert.AreEqual("0001-01-12", ws.Cell(10, 2).Value);
            Assert.AreEqual("30102017", ws.Cell(10, 3).Value);
            Assert.AreEqual("23.10.2017", ws.Cell(10, 4).Value);
            Assert.AreEqual("22.10.2017 00:00:00", ws.Cell(10, 5).Value);

            Assert.AreEqual(4d, ws.Cell(11, 1).Value);
            Assert.AreEqual(5d, ws.Cell(11, 2).Value);
            Assert.AreEqual(6d, ws.Cell(11, 3).Value);
            Assert.IsTrue(ws.Cell(11, 4).IsEmpty());
            Assert.IsTrue(ws.Cell(11, 5).IsEmpty());

            Assert.AreEqual("{p:StrParam}", ws.Cell(12, 1).Value);
            Assert.AreEqual("{m:Counter()}", ws.Cell(12, 2).Value);
            Assert.AreEqual("Plain text outside range", ws.Cell(13, 1).Value);

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());
            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);
            //report.Workbook.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestHierarchicalPanelRender()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");

            IXLRange range1 = ws.Range(1, 1, 10, 8);
            ws.Cell(1, 1).Value = "Panel1: {p:IntParam}";
            ws.Cell(10, 8).Value = "Panel1: {p:IntParam}";
            var panel1 = new ExcelPanel(range1, report);

            IXLRange range2 = ws.Range(3, 1, 8, 2);
            ws.Cell(3, 1).Value = "Panel2: {p:IntParam}";
            var panel2 = new ExcelPanel(range2, report) { Parent = panel1 };

            IXLRange range3 = ws.Range(1, 3, 6, 5);
            ws.Cell(1, 3).Value = "Panel3: {p:IntParam}";
            range3.AddToNamed("NamedPanel1");
            var panel3 = new ExcelNamedPanel(ws.Workbook.NamedRange("NamedPanel1"), report) { Parent = panel1 };

            IXLRange range4 = ws.Range(5, 6, 9, 8);
            ws.Cell(5, 6).Value = "Panel4: {p:IntParam}";
            range4.AddToNamed("NamedPanel2", XLScope.Worksheet);
            var panel4 = new ExcelNamedPanel(ws.NamedRange("NamedPanel2"), report) { Parent = panel1 };

            IXLRange range5 = ws.Range(4, 1, 5, 2);
            ws.Cell(4, 1).Value = "Panel5: {p:IntParam}";
            var panel5 = new ExcelPanel(range5, report) { Parent = panel2 };

            IXLRange range6 = ws.Range(6, 1, 8, 2);
            ws.Cell(6, 1).Value = "Panel6: {p:IntParam}";
            range6.AddToNamed("NamedPanel3");
            var panel6 = new ExcelNamedPanel(ws.Workbook.NamedRange("NamedPanel3"), report) { Parent = panel2 };

            IXLRange range7 = ws.Range(6, 1, 6, 2);
            ws.Cell(6, 2).Value = "Panel7: {p:IntParam}";
            var panel7 = new ExcelPanel(range7, report) { Parent = panel6 };

            IXLRange range8 = ws.Range(7, 1, 7, 2);
            ws.Cell(7, 2).Value = "Panel8: {p:IntParam}";
            range8.AddToNamed("NamedPanel4", XLScope.Worksheet);
            var panel8 = new ExcelNamedPanel(ws.NamedRange("NamedPanel4"), report) { Parent = panel6 };

            IXLRange range9 = ws.Range(1, 3, 6, 5);
            ws.Cell(6, 5).Value = "Panel9: {p:IntParam}";
            range9.AddToNamed("NamedPanel5", XLScope.Worksheet);
            var panel9 = new ExcelNamedPanel(ws.NamedRange("NamedPanel5"), report) { Parent = panel3 };

            IXLRange range10 = ws.Range(3, 3, 4, 5);
            ws.Cell(4, 5).Value = "Panel10: {p:IntParam}";
            var panel10 = new ExcelPanel(range10, report) { Parent = panel9 };

            IXLRange range11 = ws.Range(5, 6, 9, 8);
            ws.Cell(6, 6).Value = "Panel11: {p:IntParam}";
            var panel11 = new ExcelPanel(range11, report) { Parent = panel4 };

            IXLRange range12 = ws.Range(8, 6, 9, 8);
            ws.Cell(9, 8).Value = "Panel12: {p:IntParam}";
            range12.AddToNamed("NamedPanel6");
            var panel12 = new ExcelNamedPanel(ws.Workbook.NamedRange("NamedPanel6"), report) { Parent = panel11 };

            panel1.Children = new[] { panel2, panel3, panel4 };
            panel2.Children = new[] { panel5, panel6 };
            panel3.Children = new[] { panel9 };
            panel6.Children = new[] { panel7, panel8 };
            panel4.Children = new[] { panel11 };
            panel9.Children = new[] { panel10 };
            panel11.Children = new[] { panel12 };

            ws.Cell(11, 8).Value = "Outside panel: {p:IntParam}";

            panel1.Render();

            Assert.AreEqual(14, ws.CellsUsed().Count());
            Assert.AreEqual("Panel1: 10", ws.Cell(1, 1).Value);
            Assert.AreEqual("Panel1: 10", ws.Cell(10, 8).Value);
            Assert.AreEqual("Panel2: 10", ws.Cell(3, 1).Value);
            Assert.AreEqual("Panel3: 10", ws.Cell(1, 3).Value);
            Assert.AreEqual("Panel4: 10", ws.Cell(5, 6).Value);
            Assert.AreEqual("Panel5: 10", ws.Cell(4, 1).Value);
            Assert.AreEqual("Panel6: 10", ws.Cell(6, 1).Value);
            Assert.AreEqual("Panel7: 10", ws.Cell(6, 2).Value);
            Assert.AreEqual("Panel8: 10", ws.Cell(7, 2).Value);
            Assert.AreEqual("Panel9: 10", ws.Cell(6, 5).Value);
            Assert.AreEqual("Panel10: 10", ws.Cell(4, 5).Value);
            Assert.AreEqual("Panel11: 10", ws.Cell(6, 6).Value);
            Assert.AreEqual("Panel12: 10", ws.Cell(9, 8).Value);
            Assert.AreEqual("Outside panel: {p:IntParam}", ws.Cell(11, 8).Value);

            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());
            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);
            //report.Workbook.SaveAs("test.xlsx");
        }
    }
}