﻿using System.Reflection;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using ExcelReportGenerator.Tests.CustomAsserts;
using Assert = NUnit.Framework.Legacy.ClassicAssert;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels;

public class NamedPanelTest
{
    [Test]
    public void TestCopy()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");
        var excelReport = Substitute.For<object>();
        var templateProcessor = Substitute.For<ITemplateProcessor>();

        var range = ws.Range(1, 1, 3, 4);
        range.AddToNamed("Parent", XLScope.Worksheet);
        var namedRange = ws.DefinedName("Parent");

        var childRange = ws.Range(2, 1, 3, 4);
        childRange.AddToNamed("Child", XLScope.Worksheet);
        var namedChildRange = ws.DefinedName("Child");

        var childOfChildRange = ws.Range(3, 1, 3, 4);
        childOfChildRange.AddToNamed("ChildOfChild", XLScope.Worksheet);
        var namedChildOfChildRange = ws.DefinedName("ChildOfChild");

        var panel = new ExcelNamedPanel(namedRange, excelReport, templateProcessor)
        {
            Children = new List<IExcelPanel>
            {
                new ExcelNamedPanel(namedChildRange, excelReport, templateProcessor)
                {
                    Children = new List<IExcelPanel>
                    {
                        new ExcelDataSourcePanel("fn:DataSource:Method()", namedChildOfChildRange, excelReport,
                            templateProcessor)
                        {
                            RenderPriority = 30,
                            Type = PanelType.Horizontal,
                            ShiftType = ShiftType.Row,
                            BeforeRenderMethodName = "BeforeRenderMethod3",
                            AfterRenderMethodName = "AfterRenderMethod3",
                            BeforeDataItemRenderMethodName = "BeforeDataItemRenderMethodName",
                            AfterDataItemRenderMethodName = "AfterDataItemRenderMethodName",
                            GroupBy = "2,4"
                        }
                    },
                    RenderPriority = 20,
                    ShiftType = ShiftType.Row,
                    BeforeRenderMethodName = "BeforeRenderMethod2",
                    AfterRenderMethodName = "AfterRenderMethod2"
                }
            },
            RenderPriority = 10,
            Type = PanelType.Horizontal,
            ShiftType = ShiftType.NoShift,
            BeforeRenderMethodName = "BeforeRenderMethod1",
            AfterRenderMethodName = "AfterRenderMethod1"
        };

        var copiedPanel = (IExcelNamedPanel) panel.Copy(ws.Cell(5, 5));

        Assert.AreSame(excelReport,
            copiedPanel.GetType().GetField("_report", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(copiedPanel));
        Assert.AreSame(templateProcessor,
            copiedPanel.GetType().GetField("_templateProcessor", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(copiedPanel));
        Assert.IsTrue(Regex.IsMatch(copiedPanel.Name, @"Parent_[0-9a-f]{32}"));
        Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Range.LastCell());
        Assert.AreEqual(10, copiedPanel.RenderPriority);
        Assert.AreEqual(PanelType.Horizontal, copiedPanel.Type);
        Assert.AreEqual(ShiftType.NoShift, copiedPanel.ShiftType);
        Assert.AreEqual("BeforeRenderMethod1", copiedPanel.BeforeRenderMethodName);
        Assert.AreEqual("AfterRenderMethod1", copiedPanel.AfterRenderMethodName);
        Assert.IsNull(copiedPanel.Parent);

        Assert.AreEqual(1, copiedPanel.Children.Count);
        Assert.AreSame(excelReport,
            copiedPanel.Children.First().GetType().GetField("_report", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(copiedPanel.Children.First()));
        Assert.AreSame(templateProcessor,
            copiedPanel.Children.First().GetType()
                .GetField("_templateProcessor", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(copiedPanel.Children.First()));
        Assert.IsTrue(Regex.IsMatch(((IExcelNamedPanel) copiedPanel.Children.First()).Name,
            @"Parent_[0-9a-f]{32}_Child"));
        Assert.AreEqual(ws.Cell(6, 5), copiedPanel.Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Children.First().Range.LastCell());
        Assert.AreEqual(20, copiedPanel.Children.First().RenderPriority);
        Assert.AreEqual(PanelType.Vertical, copiedPanel.Children.First().Type);
        Assert.AreEqual(ShiftType.Row, copiedPanel.Children.First().ShiftType);
        Assert.AreEqual("BeforeRenderMethod2", copiedPanel.Children.First().BeforeRenderMethodName);
        Assert.AreEqual("AfterRenderMethod2", copiedPanel.Children.First().AfterRenderMethodName);
        Assert.AreSame(copiedPanel, copiedPanel.Children.First().Parent);

        Assert.AreEqual(1, copiedPanel.Children.First().Children.Count);
        Assert.AreSame(excelReport,
            copiedPanel.Children.First().Children.First().GetType()
                .GetField("_report", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(copiedPanel.Children.First().Children.First()));
        Assert.AreSame(templateProcessor,
            copiedPanel.Children.First().Children.First().GetType()
                .GetField("_templateProcessor", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(copiedPanel.Children.First().Children.First()));
        Assert.IsTrue(Regex.IsMatch(((IExcelNamedPanel) copiedPanel.Children.First().Children.First()).Name,
            @"Parent_[0-9a-f]{32}_Child_ChildOfChild"));
        Assert.IsInstanceOf<ExcelDataSourcePanel>(copiedPanel.Children.First().Children.First());
        Assert.AreEqual(ws.Cell(7, 5), copiedPanel.Children.First().Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Children.First().Children.First().Range.LastCell());
        Assert.AreEqual(30, copiedPanel.Children.First().Children.First().RenderPriority);
        Assert.AreEqual(PanelType.Horizontal, copiedPanel.Children.First().Children.First().Type);
        Assert.AreEqual(ShiftType.Row, copiedPanel.Children.First().Children.First().ShiftType);
        Assert.AreEqual("BeforeRenderMethod3", copiedPanel.Children.First().Children.First().BeforeRenderMethodName);
        Assert.AreEqual("AfterRenderMethod3", copiedPanel.Children.First().Children.First().AfterRenderMethodName);
        Assert.AreEqual("BeforeDataItemRenderMethodName",
            ((ExcelDataSourcePanel) copiedPanel.Children.First().Children.First()).BeforeDataItemRenderMethodName);
        Assert.AreEqual("AfterDataItemRenderMethodName",
            ((ExcelDataSourcePanel) copiedPanel.Children.First().Children.First()).AfterDataItemRenderMethodName);
        Assert.AreEqual("2,4", ((ExcelDataSourcePanel) copiedPanel.Children.First().Children.First()).GroupBy);
        Assert.AreSame(copiedPanel.Children.First(), copiedPanel.Children.First().Children.First().Parent);

        namedRange.Delete();
        namedChildRange.Delete();
        copiedPanel.Delete();
        copiedPanel.Children.First().Delete();

        IExcelPanel globalParent = new ExcelPanel(ws.Range(1, 1, 20, 20), excelReport, templateProcessor);
        range = ws.Range(1, 1, 3, 4);
        range.AddToNamed("Parent", XLScope.Worksheet);
        namedRange = ws.DefinedName("Parent");

        var childRange1 = ws.Range(1, 1, 1, 4);
        childRange1.AddToNamed("Child", XLScope.Worksheet);
        namedChildRange = ws.DefinedName("Child");

        var childRange2 = ws.Range(2, 1, 3, 4);

        childOfChildRange = ws.Range(3, 1, 3, 4);
        childOfChildRange.AddToNamed("ChildOfChild", XLScope.Worksheet);
        namedChildOfChildRange = ws.DefinedName("ChildOfChild");

        panel = new ExcelNamedPanel(namedRange, excelReport, templateProcessor)
        {
            Parent = globalParent,
            Children = new List<IExcelPanel>
            {
                new ExcelNamedPanel(namedChildRange, excelReport, templateProcessor),
                new ExcelPanel(childRange2, excelReport, templateProcessor)
                {
                    Children = new List<IExcelPanel>
                    {
                        new ExcelNamedPanel(namedChildOfChildRange, excelReport, templateProcessor)
                    },
                    RenderPriority = 10,
                    Type = PanelType.Horizontal,
                    ShiftType = ShiftType.NoShift,
                    BeforeRenderMethodName = "BeforeRenderMethod",
                    AfterRenderMethodName = "AfterRenderMethod"
                }
            }
        };

        copiedPanel = (IExcelNamedPanel) panel.Copy(ws.Cell(5, 5));
        Assert.IsTrue(Regex.IsMatch(copiedPanel.Name, @"Parent_[0-9a-f]{32}"));
        Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Range.LastCell());
        Assert.AreSame(globalParent, copiedPanel.Parent);

        Assert.AreEqual(2, copiedPanel.Children.Count);
        Assert.IsTrue(Regex.IsMatch(((IExcelNamedPanel) copiedPanel.Children.First()).Name,
            @"Parent_[0-9a-f]{32}_Child"));
        Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(5, 8), copiedPanel.Children.First().Range.LastCell());
        Assert.AreSame(copiedPanel, copiedPanel.Children.First().Parent);
        Assert.IsInstanceOf<ExcelPanel>(copiedPanel.Children.Last());
        Assert.IsNotInstanceOf<ExcelNamedPanel>(copiedPanel.Children.Last());
        Assert.AreEqual(ws.Cell(6, 5), copiedPanel.Children.Last().Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Children.Last().Range.LastCell());
        Assert.AreEqual(10, copiedPanel.Children.Last().RenderPriority);
        Assert.AreEqual(PanelType.Horizontal, copiedPanel.Children.Last().Type);
        Assert.AreEqual(ShiftType.NoShift, copiedPanel.Children.Last().ShiftType);
        Assert.AreEqual("BeforeRenderMethod", copiedPanel.Children.Last().BeforeRenderMethodName);
        Assert.AreEqual("AfterRenderMethod", copiedPanel.Children.Last().AfterRenderMethodName);

        Assert.AreSame(copiedPanel, copiedPanel.Children.Last().Parent);

        Assert.AreEqual(1, copiedPanel.Children.Last().Children.Count);
        Assert.IsTrue(Regex.IsMatch(((IExcelNamedPanel) copiedPanel.Children.Last().Children.First()).Name,
            @"ChildOfChild_[0-9a-f]{32}"));
        Assert.AreEqual(ws.Cell(7, 5), copiedPanel.Children.Last().Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Children.Last().Children.First().Range.LastCell());
        Assert.AreSame(copiedPanel.Children.Last(), copiedPanel.Children.Last().Children.First().Parent);

        namedRange.Delete();
        namedChildRange.Delete();
        namedChildOfChildRange.Delete();
        copiedPanel.Delete();
        copiedPanel.Children.First().Delete();
        copiedPanel.Children.Last().Children.First().Delete();

        globalParent = new ExcelPanel(ws.Range(1, 1, 7, 7), excelReport, templateProcessor);
        range = ws.Range(1, 1, 3, 4);
        range.AddToNamed("Parent", XLScope.Worksheet);
        namedRange = ws.DefinedName("Parent");

        childRange = ws.Range(1, 1, 1, 4);
        childRange.AddToNamed("Child", XLScope.Worksheet);
        namedChildRange = ws.DefinedName("Child");

        panel = new ExcelNamedPanel(namedRange, excelReport, templateProcessor)
        {
            Parent = globalParent,
            Children = new List<IExcelPanel> {new ExcelNamedPanel(namedChildRange, excelReport, templateProcessor)}
        };

        copiedPanel = (IExcelNamedPanel) panel.Copy(ws.Cell(5, 5));
        Assert.IsTrue(Regex.IsMatch(copiedPanel.Name, @"Parent_[0-9a-f]{32}"));
        Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Range.LastCell());
        Assert.IsNull(copiedPanel.Parent);

        Assert.AreEqual(1, copiedPanel.Children.Count);
        Assert.IsTrue(Regex.IsMatch(((IExcelNamedPanel) copiedPanel.Children.First()).Name,
            @"Parent_[0-9a-f]{32}_Child"));
        Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(5, 8), copiedPanel.Children.First().Range.LastCell());
        Assert.AreSame(copiedPanel, copiedPanel.Children.First().Parent);

        copiedPanel = (IExcelNamedPanel) panel.Copy(ws.Cell(5, 5), false);
        Assert.IsTrue(Regex.IsMatch(copiedPanel.Name, @"Parent_[0-9a-f]{32}"));
        Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Range.LastCell());
        Assert.IsNull(copiedPanel.Parent);
        Assert.AreEqual(0, copiedPanel.Children.Count);

        ExceptionAssert.Throws<ArgumentNullException>(() => panel.Copy(null));

        //wb.SaveAs("test.xlsx");
    }

    [Test]
    public void TestCopyWithName()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");
        var excelReport = Substitute.For<object>();
        var templateProcessor = Substitute.For<ITemplateProcessor>();

        var range = ws.Range(1, 1, 3, 4);
        range.AddToNamed("Parent", XLScope.Worksheet);
        var namedRange = ws.DefinedName("Parent");

        var childRange1 = ws.Range(1, 1, 2, 4);
        childRange1.AddToNamed("Child", XLScope.Worksheet);
        var namedChildRange = ws.DefinedName("Child");

        var childRange2 = ws.Range(3, 1, 3, 4);

        var childOfChildRange1 = ws.Range(1, 1, 1, 4);
        childOfChildRange1.AddToNamed("ChildOfChild1", XLScope.Worksheet);
        var namedChildOfChildRange1 = ws.DefinedName("ChildOfChild1");

        var childOfChildRange2 = ws.Range(3, 1, 3, 4);
        childOfChildRange2.AddToNamed("ChildOfChild2", XLScope.Worksheet);
        var namedChildOfChildRange2 = ws.DefinedName("ChildOfChild2");

        var panel = new ExcelNamedPanel(namedRange, excelReport, templateProcessor)
        {
            Children = new List<IExcelPanel>
            {
                new ExcelNamedPanel(namedChildRange, excelReport, templateProcessor)
                {
                    Children = new List<IExcelPanel>
                    {
                        new ExcelNamedPanel(namedChildOfChildRange1, excelReport, templateProcessor)
                    }
                },
                new ExcelPanel(childRange2, excelReport, templateProcessor)
                {
                    Children = new List<IExcelPanel>
                    {
                        new ExcelNamedPanel(namedChildOfChildRange2, excelReport, templateProcessor)
                    }
                }
            }
        };

        var copiedPanel = panel.Copy(ws.Cell(5, 5), "Copied");
        Assert.IsTrue(Regex.IsMatch(copiedPanel.Name, "Copied"));
        Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Range.LastCell());
        Assert.IsNull(copiedPanel.Parent);

        Assert.AreEqual(2, copiedPanel.Children.Count);
        Assert.IsTrue(Regex.IsMatch(((IExcelNamedPanel) copiedPanel.Children.First()).Name, "Copied_Child"));
        Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(6, 8), copiedPanel.Children.First().Range.LastCell());
        Assert.AreSame(copiedPanel, copiedPanel.Children.First().Parent);
        Assert.IsInstanceOf<ExcelPanel>(copiedPanel.Children.Last());
        Assert.IsNotInstanceOf<ExcelNamedPanel>(copiedPanel.Children.Last());
        Assert.AreEqual(ws.Cell(7, 5), copiedPanel.Children.Last().Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Children.Last().Range.LastCell());
        Assert.AreSame(copiedPanel, copiedPanel.Children.Last().Parent);

        Assert.AreEqual(1, copiedPanel.Children.First().Children.Count);
        Assert.IsTrue(Regex.IsMatch(((IExcelNamedPanel) copiedPanel.Children.First().Children.First()).Name,
            "Copied_Child_ChildOfChild1"));
        Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Children.First().Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(5, 8), copiedPanel.Children.First().Children.First().Range.LastCell());
        Assert.AreSame(copiedPanel.Children.First(), copiedPanel.Children.First().Children.First().Parent);

        Assert.AreEqual(1, copiedPanel.Children.Last().Children.Count);
        Assert.IsTrue(Regex.IsMatch(((IExcelNamedPanel) copiedPanel.Children.Last().Children.First()).Name,
            "ChildOfChild2_[0-9a-f]{32}"));
        Assert.AreEqual(ws.Cell(7, 5), copiedPanel.Children.Last().Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Children.Last().Children.First().Range.LastCell());
        Assert.AreSame(copiedPanel.Children.Last(), copiedPanel.Children.Last().Children.First().Parent);

        copiedPanel = panel.Copy(ws.Cell(5, 5), "Copied2", false);
        Assert.IsTrue(Regex.IsMatch(copiedPanel.Name, "Copied2"));
        Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Range.LastCell());
        Assert.IsNull(copiedPanel.Parent);
        Assert.AreEqual(0, copiedPanel.Children.Count);

        ExceptionAssert.Throws<ArgumentNullException>(() => panel.Copy(null, "Copied"));
        ExceptionAssert.Throws<ArgumentException>(() => panel.Copy(ws.Cell(5, 5), null));
        ExceptionAssert.Throws<ArgumentException>(() => panel.Copy(ws.Cell(5, 5), string.Empty));
        ExceptionAssert.Throws<ArgumentException>(() => panel.Copy(ws.Cell(5, 5), " "));

        //wb.SaveAs("test.xlsx");
    }

    [Test]
    public void TestMove()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");
        var excelReport = Substitute.For<object>();
        var templateProcessor = Substitute.For<ITemplateProcessor>();

        var range = ws.Range(1, 1, 4, 5);
        range.AddToNamed("parentRange", XLScope.Worksheet);
        var namedParentRange = ws.DefinedName("parentRange");

        var childRange1 = ws.Range(1, 1, 2, 5);
        var childRange2 = ws.Range(3, 1, 4, 5);
        childRange2.AddToNamed("childRange2", XLScope.Worksheet);
        var namedChildRange = ws.DefinedName("childRange2");

        var childOfChildRange1 = ws.Range(2, 1, 2, 5);
        childOfChildRange1.AddToNamed("childOfChildRange1", XLScope.Worksheet);
        var childOfChildNamedRange = ws.DefinedName("childOfChildRange1");

        var childOfChildRange2 = ws.Range(4, 1, 4, 5);

        var panel = new ExcelNamedPanel(namedParentRange, excelReport, templateProcessor)
        {
            Children = new List<IExcelPanel>
            {
                new ExcelPanel(childRange1, excelReport, templateProcessor)
                {
                    Children = new List<IExcelPanel>
                    {
                        new ExcelDataSourcePanel("fn:DataSource:Method()", childOfChildNamedRange, excelReport,
                            templateProcessor)
                    }
                },
                new ExcelNamedPanel(namedChildRange, excelReport, templateProcessor)
                {
                    Children = new List<IExcelPanel>
                    {
                        new ExcelPanel(childOfChildRange2, excelReport, templateProcessor)
                    }
                }
            }
        };

        IExcelPanel globalParent = new ExcelPanel(ws.Range(1, 1, 8, 10), excelReport, templateProcessor);

        panel.Children.First().Children.First().Parent = panel.Children.First();
        panel.Children.Last().Children.First().Parent = panel.Children.Last();
        panel.Children.ToList().ForEach(c => c.Parent = panel);
        panel.Parent = globalParent;

        panel.Move(ws.Cell(5, 6));

        Assert.AreEqual(ws.Cell(5, 6), panel.Range.FirstCell());
        Assert.AreEqual(ws.Cell(8, 10), panel.Range.LastCell());
        Assert.AreEqual("parentRange", ((IExcelNamedPanel) panel).Name);
        Assert.AreSame(globalParent, panel.Parent);

        Assert.AreEqual(2, panel.Children.Count());
        Assert.AreEqual(ws.Cell(5, 6), panel.Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(6, 10), panel.Children.First().Range.LastCell());
        Assert.IsInstanceOf<ExcelPanel>(panel.Children.First());
        Assert.IsNotInstanceOf<IExcelNamedPanel>(panel.Children.First());
        Assert.AreSame(panel, panel.Children.First().Parent);

        Assert.AreEqual(ws.Cell(7, 6), panel.Children.Last().Range.FirstCell());
        Assert.AreEqual(ws.Cell(8, 10), panel.Children.Last().Range.LastCell());
        Assert.AreEqual("childRange2", ((IExcelNamedPanel) panel.Children.Last()).Name);
        Assert.AreSame(panel, panel.Children.First().Parent);

        Assert.AreEqual(1, panel.Children.First().Children.Count());
        Assert.AreEqual(ws.Cell(6, 6), panel.Children.First().Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(6, 10), panel.Children.First().Children.First().Range.LastCell());
        Assert.IsInstanceOf<ExcelDataSourcePanel>(panel.Children.First().Children.First());
        Assert.AreEqual("childOfChildRange1", ((IExcelNamedPanel) panel.Children.First().Children.First()).Name);
        Assert.AreSame(panel.Children.First(), panel.Children.First().Children.First().Parent);

        Assert.AreEqual(1, panel.Children.Last().Children.Count());
        Assert.AreEqual(ws.Cell(8, 6), panel.Children.Last().Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(8, 10), panel.Children.Last().Children.First().Range.LastCell());
        Assert.IsInstanceOf<ExcelPanel>(panel.Children.Last().Children.First());
        Assert.IsNotInstanceOf<IExcelNamedPanel>(panel.Children.Last().Children.First());
        Assert.AreSame(panel.Children.Last(), panel.Children.Last().Children.First().Parent);

        Assert.AreEqual(3, ws.DefinedNames.Count());

        panel.Move(ws.Cell(6, 6));

        Assert.AreEqual(ws.Cell(6, 6), panel.Range.FirstCell());
        Assert.AreEqual(ws.Cell(9, 10), panel.Range.LastCell());
        Assert.IsNull(panel.Parent);

        Assert.AreEqual(2, panel.Children.Count());
        Assert.AreEqual(ws.Cell(6, 6), panel.Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 10), panel.Children.First().Range.LastCell());
        Assert.IsInstanceOf<ExcelPanel>(panel.Children.First());
        Assert.IsNotInstanceOf<IExcelNamedPanel>(panel.Children.First());
        Assert.AreSame(panel, panel.Children.First().Parent);

        Assert.AreEqual(ws.Cell(8, 6), panel.Children.Last().Range.FirstCell());
        Assert.AreEqual(ws.Cell(9, 10), panel.Children.Last().Range.LastCell());
        Assert.AreEqual("childRange2", ((IExcelNamedPanel) panel.Children.Last()).Name);
        Assert.AreSame(panel, panel.Children.First().Parent);

        Assert.AreEqual(1, panel.Children.First().Children.Count());
        Assert.AreEqual(ws.Cell(7, 6), panel.Children.First().Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 10), panel.Children.First().Children.First().Range.LastCell());
        Assert.IsInstanceOf<ExcelDataSourcePanel>(panel.Children.First().Children.First());
        Assert.AreEqual("childOfChildRange1", ((IExcelNamedPanel) panel.Children.First().Children.First()).Name);
        Assert.AreSame(panel.Children.First(), panel.Children.First().Children.First().Parent);

        Assert.AreEqual(1, panel.Children.Last().Children.Count());
        Assert.AreEqual(ws.Cell(9, 6), panel.Children.Last().Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(9, 10), panel.Children.Last().Children.First().Range.LastCell());
        Assert.IsInstanceOf<ExcelPanel>(panel.Children.Last().Children.First());
        Assert.IsNotInstanceOf<IExcelNamedPanel>(panel.Children.Last().Children.First());
        Assert.AreSame(panel.Children.Last(), panel.Children.Last().Children.First().Parent);

        Assert.AreEqual(3, ws.DefinedNames.Count());

        //wb.SaveAs("test.xlsx");
    }

    [Test]
    public void TestGetNearestNamedPanelTest()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");
        var excelReport = Substitute.For<object>();
        var templateProcessor = Substitute.For<ITemplateProcessor>();

        var range = ws.Range(1, 1, 3, 4);
        range.AddToNamed("Parent", XLScope.Worksheet);
        var namedRange = ws.DefinedName("Parent");

        var childRange = ws.Range(2, 1, 3, 4);

        var childOfChildRange = ws.Range(3, 1, 3, 4);
        childOfChildRange.AddToNamed("ChildOfChild", XLScope.Worksheet);
        var childOfChildNamedRange = ws.DefinedName("ChildOfChild");

        IExcelPanel childOfChildPanel = new ExcelNamedPanel(childOfChildNamedRange, excelReport, templateProcessor);
        IExcelPanel childPanel = new ExcelPanel(childRange, excelReport, templateProcessor);
        IExcelPanel parentPanel = new ExcelNamedPanel(namedRange, excelReport, templateProcessor);

        var method =
            typeof(ExcelNamedPanel).GetMethod("GetNearestNamedParent", BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.IsNull(method.Invoke(childOfChildPanel, null));

        childOfChildPanel.Parent = childPanel;
        Assert.IsNull(method.Invoke(childOfChildPanel, null));

        childPanel.Parent = parentPanel;
        Assert.AreSame(parentPanel, method.Invoke(childOfChildPanel, null));
    }

    [Test]
    public void TestRemoveName()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");
        var excelReport = Substitute.For<object>();
        var templateProcessor = Substitute.For<ITemplateProcessor>();

        var range = ws.Range(1, 1, 3, 4);
        range.AddToNamed("Parent", XLScope.Worksheet);
        var namedRange = ws.DefinedName("Parent");
        IExcelNamedPanel parentPanel = new ExcelNamedPanel(namedRange, excelReport, templateProcessor);

        var childRange1 = ws.Range(1, 1, 1, 4);
        childRange1.AddToNamed("Child", XLScope.Worksheet);
        var namedChildRange = ws.DefinedName("Child");
        IExcelNamedPanel childPanel1 = new ExcelNamedPanel(namedChildRange, excelReport, templateProcessor);
        childPanel1.Parent = parentPanel;

        var childRange2 = ws.Range(2, 1, 3, 4);
        IExcelPanel childPanel2 = new ExcelPanel(childRange2, excelReport, templateProcessor);
        childPanel2.Parent = parentPanel;

        parentPanel.Children = new List<IExcelPanel> {childPanel1, childPanel2};

        var childOfChild1Range = ws.Range(1, 1, 1, 4);
        childOfChild1Range.AddToNamed("ChildOfChild1", XLScope.Worksheet);
        var namedChildOfChild1RangeRange = ws.DefinedName("ChildOfChild1");
        IExcelNamedPanel childOfChild1Panel =
            new ExcelNamedPanel(namedChildOfChild1RangeRange, excelReport, templateProcessor);
        childOfChild1Panel.Parent = childPanel1;
        childPanel1.Children = new List<IExcelPanel> {childOfChild1Panel};

        var childOfChild2Range = ws.Range(3, 1, 3, 4);
        childOfChild2Range.AddToNamed("ChildOfChild2", XLScope.Worksheet);
        var namedChildOfChild2RangeRange = ws.DefinedName("ChildOfChild2");
        IExcelNamedPanel childOfChild2Panel =
            new ExcelNamedPanel(namedChildOfChild2RangeRange, excelReport, templateProcessor);
        childOfChild2Panel.Parent = childPanel2;
        childPanel2.Children = new List<IExcelPanel> {childOfChild2Panel};

        parentPanel.RemoveName();
        Assert.AreEqual(3, ws.DefinedNames.Count());
        Assert.IsNull(ws.DefinedNames.SingleOrDefault(r => r.Name == "Parent"));

        range.AddToNamed("Parent", XLScope.Worksheet);
        Assert.AreEqual(4, ws.DefinedNames.Count());
        Assert.IsNotNull(ws.DefinedNames.SingleOrDefault(r => r.Name == "Parent"));

        parentPanel.RemoveName(true);
        Assert.AreEqual(0, ws.DefinedNames.Count());

        //wb.SaveAs("test.xlsx");
    }

    [Test]
    public void TestRemoveAllNamesRecursive()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");
        var excelReport = Substitute.For<object>();
        var templateProcessor = Substitute.For<ITemplateProcessor>();

        var range = ws.Range(1, 1, 3, 4);
        range.AddToNamed("Parent", XLScope.Worksheet);
        var namedRange = ws.DefinedName("Parent");
        IExcelNamedPanel parentPanel = new ExcelNamedPanel(namedRange, excelReport, templateProcessor);

        var childRange1 = ws.Range(1, 1, 1, 4);
        childRange1.AddToNamed("Child", XLScope.Worksheet);
        var namedChildRange = ws.DefinedName("Child");
        IExcelNamedPanel childPanel1 = new ExcelNamedPanel(namedChildRange, excelReport, templateProcessor);
        childPanel1.Parent = parentPanel;

        var childRange2 = ws.Range(2, 1, 3, 4);
        IExcelPanel childPanel2 = new ExcelPanel(childRange2, excelReport, templateProcessor);
        childPanel2.Parent = parentPanel;

        parentPanel.Children = new List<IExcelPanel> {childPanel1, childPanel2};

        var childOfChild1Range = ws.Range(1, 1, 1, 4);
        childOfChild1Range.AddToNamed("ChildOfChild1", XLScope.Worksheet);
        var namedChildOfChild1RangeRange = ws.DefinedName("ChildOfChild1");
        IExcelNamedPanel childOfChild1Panel =
            new ExcelNamedPanel(namedChildOfChild1RangeRange, excelReport, templateProcessor);
        childOfChild1Panel.Parent = childPanel1;
        childPanel1.Children = new List<IExcelPanel> {childOfChild1Panel};

        var childOfChild2Range = ws.Range(3, 1, 3, 4);
        childOfChild2Range.AddToNamed("ChildOfChild2", XLScope.Worksheet);
        var namedChildOfChild2RangeRange = ws.DefinedName("ChildOfChild2");
        IExcelNamedPanel childOfChild2Panel =
            new ExcelNamedPanel(namedChildOfChild2RangeRange, excelReport, templateProcessor);
        childOfChild2Panel.Parent = childPanel2;
        childPanel2.Children = new List<IExcelPanel> {childOfChild2Panel};

        ExcelNamedPanel.RemoveAllNamesRecursive(parentPanel);
        Assert.AreEqual(0, ws.DefinedNames.Count());

        //wb.SaveAs("test.xlsx");
    }

    [Test]
    public void TestDelete()
    {
        // Deleting with moving cells up
        var wb = InitWorkBookForDeleteRangeTest();
        var ws = wb.Worksheet("Test");
        var parentRange = ws.DefinedName("Parent");
        var childRange = ws.DefinedName("Child");
        Assert.AreEqual(2, ws.DefinedNames.Count());
        var excelReport = Substitute.For<object>();
        var templateProcessor = Substitute.For<ITemplateProcessor>();

        var panel = new ExcelNamedPanel(parentRange, excelReport, templateProcessor)
        {
            Children = new List<IExcelPanel> {new ExcelNamedPanel(childRange, excelReport, templateProcessor)}
        };
        panel.Delete();

        var rangeStartCell = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RangeStart");
        var rangeEndCell = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RangeEnd");
        var belowCell1 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_1");
        var belowCell2 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_2");
        var rightCell1 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_1");
        var rightCell2 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_2");
        var aboveCell1 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_1");
        var aboveCell2 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_2");
        var leftCell1 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_1");
        var leftCell2 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_2");

        Assert.IsNull(rangeStartCell);
        Assert.IsNull(rangeEndCell);
        Assert.AreEqual(8, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
        Assert.AreEqual(belowCell1, ws.Cell(6, 6));
        Assert.AreEqual(belowCell2, ws.Cell(10, 8));
        Assert.AreEqual(rightCell1, ws.Cell(7, 8));
        Assert.AreEqual(rightCell2, ws.Cell(5, 8));
        Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
        Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
        Assert.AreEqual(leftCell1, ws.Cell(7, 4));
        Assert.AreEqual(leftCell2, ws.Cell(10, 4));
        Assert.AreEqual(0, ws.DefinedNames.Count());

        // Deleting with moving the row up
        wb = InitWorkBookForDeleteRangeTest();
        ws = wb.Worksheet("Test");
        parentRange = ws.DefinedName("Parent");
        childRange = ws.DefinedName("Child");
        Assert.AreEqual(2, ws.DefinedNames.Count());

        panel = new ExcelNamedPanel(parentRange, excelReport, templateProcessor)
        {
            Children = new List<IExcelPanel> {new ExcelNamedPanel(childRange, excelReport, templateProcessor)},
            ShiftType = ShiftType.Row
        };
        panel.Delete();

        rangeStartCell = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RangeStart");
        rangeEndCell = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RangeEnd");
        belowCell1 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_1");
        belowCell2 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_2");
        rightCell1 = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RightCell_1");
        rightCell2 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_2");
        aboveCell1 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_1");
        aboveCell2 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_2");
        leftCell1 = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "LeftCell_1");
        leftCell2 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_2");

        Assert.IsNull(rangeStartCell);
        Assert.IsNull(rangeEndCell);
        Assert.IsNull(leftCell1);
        Assert.IsNull(rightCell1);
        Assert.AreEqual(6, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
        Assert.AreEqual(belowCell1, ws.Cell(6, 6));
        Assert.AreEqual(belowCell2, ws.Cell(6, 8));
        Assert.AreEqual(rightCell2, ws.Cell(5, 8));
        Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
        Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
        Assert.AreEqual(leftCell2, ws.Cell(6, 4));
        Assert.AreEqual(0, ws.DefinedNames.Count());

        // Deleting with moving cells left
        wb = InitWorkBookForDeleteRangeTest();
        ws = wb.Worksheet("Test");
        parentRange = ws.DefinedName("Parent");
        childRange = ws.DefinedName("Child");
        Assert.AreEqual(2, ws.DefinedNames.Count());

        panel = new ExcelNamedPanel(parentRange, excelReport, templateProcessor)
        {
            Children = new List<IExcelPanel> {new ExcelNamedPanel(childRange, excelReport, templateProcessor)},
            Type = PanelType.Horizontal
        };
        panel.Delete();

        rangeStartCell = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RangeStart");
        rangeEndCell = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RangeEnd");
        belowCell1 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_1");
        belowCell2 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_2");
        rightCell1 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_1");
        rightCell2 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_2");
        aboveCell1 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_1");
        aboveCell2 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_2");
        leftCell1 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_1");
        leftCell2 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_2");

        Assert.IsNull(rangeStartCell);
        Assert.IsNull(rangeEndCell);
        Assert.AreEqual(8, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
        Assert.AreEqual(belowCell1, ws.Cell(10, 6));
        Assert.AreEqual(belowCell2, ws.Cell(10, 8));
        Assert.AreEqual(rightCell1, ws.Cell(7, 5));
        Assert.AreEqual(rightCell2, ws.Cell(5, 8));
        Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
        Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
        Assert.AreEqual(leftCell1, ws.Cell(7, 4));
        Assert.AreEqual(leftCell2, ws.Cell(10, 4));
        Assert.AreEqual(0, ws.DefinedNames.Count());

        // Deleting with moving the column left
        wb = InitWorkBookForDeleteRangeTest();
        ws = wb.Worksheet("Test");
        parentRange = ws.DefinedName("Parent");
        childRange = ws.DefinedName("Child");
        Assert.AreEqual(2, ws.DefinedNames.Count());

        panel = new ExcelNamedPanel(parentRange, excelReport, templateProcessor)
        {
            Children = new List<IExcelPanel> {new ExcelNamedPanel(childRange, excelReport, templateProcessor)},
            Type = PanelType.Horizontal,
            ShiftType = ShiftType.Row
        };
        panel.Delete();

        rangeStartCell = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RangeStart");
        rangeEndCell = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RangeEnd");
        belowCell1 = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "BelowCell_1");
        belowCell2 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_2");
        rightCell1 = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RightCell_1");
        rightCell2 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_2");
        aboveCell1 = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "AboveCell_1");
        aboveCell2 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_2");
        leftCell1 = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "LeftCell_1");
        leftCell2 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_2");

        Assert.IsNull(rangeStartCell);
        Assert.IsNull(rangeEndCell);
        Assert.IsNull(aboveCell1);
        Assert.IsNull(belowCell1);
        Assert.AreEqual(6, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
        Assert.AreEqual(belowCell2, ws.Cell(10, 5));
        Assert.AreEqual(rightCell1, ws.Cell(7, 5));
        Assert.AreEqual(rightCell2, ws.Cell(5, 5));
        Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
        Assert.AreEqual(leftCell1, ws.Cell(7, 4));
        Assert.AreEqual(leftCell2, ws.Cell(10, 4));
        Assert.AreEqual(0, ws.DefinedNames.Count());

        // Deleting without any shift
        wb = InitWorkBookForDeleteRangeTest();
        ws = wb.Worksheet("Test");
        parentRange = ws.DefinedName("Parent");
        childRange = ws.DefinedName("Child");
        Assert.AreEqual(2, ws.DefinedNames.Count());

        panel = new ExcelNamedPanel(parentRange, excelReport, templateProcessor)
        {
            Children = new List<IExcelPanel> {new ExcelNamedPanel(childRange, excelReport, templateProcessor)},
            ShiftType = ShiftType.NoShift
        };
        panel.Delete();

        rangeStartCell = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RangeStart");
        rangeEndCell = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RangeEnd");
        belowCell1 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_1");
        belowCell2 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_2");
        rightCell1 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_1");
        rightCell2 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_2");
        aboveCell1 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_1");
        aboveCell2 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_2");
        leftCell1 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_1");
        leftCell2 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_2");

        Assert.IsNull(rangeStartCell);
        Assert.IsNull(rangeEndCell);
        Assert.AreEqual(XLBorderStyleValues.None, parentRange.Ranges.ElementAt(0).FirstCell().Style.Border.TopBorder);
        Assert.AreEqual(XLBorderStyleValues.None, parentRange.Ranges.ElementAt(0).Style.Border.BottomBorder);
        Assert.AreEqual(8, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
        Assert.AreEqual(belowCell1, ws.Cell(10, 6));
        Assert.AreEqual(belowCell2, ws.Cell(10, 8));
        Assert.AreEqual(rightCell1, ws.Cell(7, 8));
        Assert.AreEqual(rightCell2, ws.Cell(5, 8));
        Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
        Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
        Assert.AreEqual(leftCell1, ws.Cell(7, 4));
        Assert.AreEqual(leftCell2, ws.Cell(10, 4));
        Assert.AreEqual(0, ws.DefinedNames.Count());

        //wb.SaveAs("test.xlsx");
    }

    private XLWorkbook InitWorkBookForDeleteRangeTest()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");

        var range = ws.Range(6, 5, 9, 7);
        range.AddToNamed("Parent", XLScope.Worksheet);
        range.FirstCell().Value = "RangeStart";
        range.LastCell().Value = "RangeEnd";
        range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
        range.LastCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);

        var childRange = ws.Range(7, 5, 8, 7);
        childRange.AddToNamed("Child", XLScope.Worksheet);

        ws.Cell(10, 6).Value = "BelowCell_1";
        ws.Cell(10, 8).Value = "BelowCell_2";
        ws.Cell(7, 8).Value = "RightCell_1";
        ws.Cell(5, 8).Value = "RightCell_2";
        ws.Cell(5, 6).Value = "AboveCell_1";
        ws.Cell(5, 4).Value = "AboveCell_2";
        ws.Cell(7, 4).Value = "LeftCell_1";
        ws.Cell(10, 4).Value = "LeftCell_2";

        return wb;
    }
}