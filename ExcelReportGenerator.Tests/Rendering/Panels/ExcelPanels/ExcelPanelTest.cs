using System.Reflection;
using ClosedXML.Excel;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Exceptions;
using ExcelReportGenerator.Rendering.EventArgs;
using ExcelReportGenerator.Rendering.Panels;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using ExcelReportGenerator.Tests.CustomAsserts;
using Assert = NUnit.Framework.Legacy.ClassicAssert;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels;

public class PanelTest
{
    [Test]
    public void TestCopy()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");
        var excelReport = Substitute.For<object>();
        var templateProcessor = Substitute.For<ITemplateProcessor>();

        var range = ws.Range(1, 1, 3, 4);
        var childRange = ws.Range(2, 1, 3, 4);
        var childOfChildRange = ws.Range(3, 1, 3, 4);

        var panel = new ExcelPanel(range, excelReport, templateProcessor)
        {
            BeforeRenderMethodName = "BeforeMethod",
            AfterRenderMethodName = "AfterRender",
            Type = PanelType.Horizontal,
            ShiftType = ShiftType.Row,
            RenderPriority = 1,
            Children = new List<IExcelPanel>
            {
                new ExcelPanel(childRange, excelReport, templateProcessor)
                {
                    BeforeRenderMethodName = "BeforeMethod_child",
                    AfterRenderMethodName = "AfterRender_child",
                    Type = PanelType.Vertical,
                    ShiftType = ShiftType.NoShift,
                    RenderPriority = 2,
                    Children = new List<IExcelPanel>
                    {
                        new ExcelPanel(childOfChildRange, excelReport, templateProcessor)
                        {
                            BeforeRenderMethodName = "BeforeMethod_child_child",
                            AfterRenderMethodName = "AfterRender_child_child",
                            Type = PanelType.Horizontal,
                            ShiftType = ShiftType.Row,
                            RenderPriority = 3
                        }
                    }
                }
            }
        };

        var copiedPanel = panel.Copy(ws.Cell(5, 5));

        Assert.AreSame(excelReport,
            copiedPanel.GetType().GetField("_report", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(copiedPanel));
        Assert.AreSame(templateProcessor,
            copiedPanel.GetType().GetField("_templateProcessor", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(copiedPanel));
        Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Range.LastCell());
        Assert.IsNull(copiedPanel.Parent);
        Assert.AreEqual(panel.BeforeRenderMethodName, copiedPanel.BeforeRenderMethodName);
        Assert.AreEqual(panel.AfterRenderMethodName, copiedPanel.AfterRenderMethodName);
        Assert.AreEqual(panel.Type, copiedPanel.Type);
        Assert.AreEqual(panel.ShiftType, copiedPanel.ShiftType);
        Assert.AreEqual(panel.RenderPriority, copiedPanel.RenderPriority);

        Assert.AreEqual(1, copiedPanel.Children.Count());
        Assert.AreSame(excelReport,
            copiedPanel.Children.First().GetType().GetField("_report", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(copiedPanel.Children.First()));
        Assert.AreSame(templateProcessor,
            copiedPanel.GetType().GetField("_templateProcessor", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(copiedPanel.Children.First()));
        Assert.AreEqual(ws.Cell(6, 5), copiedPanel.Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Children.First().Range.LastCell());
        Assert.AreSame(copiedPanel, copiedPanel.Children.First().Parent);
        Assert.AreEqual(panel.Children.First().BeforeRenderMethodName,
            copiedPanel.Children.First().BeforeRenderMethodName);
        Assert.AreEqual(panel.Children.First().AfterRenderMethodName,
            copiedPanel.Children.First().AfterRenderMethodName);
        Assert.AreEqual(panel.Children.First().Type, copiedPanel.Children.First().Type);
        Assert.AreEqual(panel.Children.First().ShiftType, copiedPanel.Children.First().ShiftType);
        Assert.AreEqual(panel.Children.First().RenderPriority, copiedPanel.Children.First().RenderPriority);

        Assert.AreEqual(1, copiedPanel.Children.First().Children.Count());
        Assert.AreSame(excelReport,
            copiedPanel.Children.First().Children.First().GetType()
                .GetField("_report", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(copiedPanel.Children.First().Children.First()));
        Assert.AreSame(templateProcessor,
            copiedPanel.Children.First().Children.First().GetType()
                .GetField("_templateProcessor", BindingFlags.Instance | BindingFlags.NonPublic)
                .GetValue(copiedPanel.Children.First().Children.First()));
        Assert.AreEqual(ws.Cell(7, 5), copiedPanel.Children.First().Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Children.First().Children.First().Range.LastCell());
        Assert.AreSame(copiedPanel.Children.First(), copiedPanel.Children.First().Children.First().Parent);
        Assert.AreEqual(panel.Children.First().Children.First().BeforeRenderMethodName,
            copiedPanel.Children.First().Children.First().BeforeRenderMethodName);
        Assert.AreEqual(panel.Children.First().Children.First().AfterRenderMethodName,
            copiedPanel.Children.First().Children.First().AfterRenderMethodName);
        Assert.AreEqual(panel.Children.First().Children.First().Type,
            copiedPanel.Children.First().Children.First().Type);
        Assert.AreEqual(panel.Children.First().Children.First().ShiftType,
            copiedPanel.Children.First().Children.First().ShiftType);
        Assert.AreEqual(panel.Children.First().Children.First().RenderPriority,
            copiedPanel.Children.First().Children.First().RenderPriority);

        IExcelPanel globalParent = new ExcelPanel(ws.Range(1, 1, 20, 20), excelReport, templateProcessor);
        range = ws.Range(1, 1, 3, 4);
        var childRange1 = ws.Range(1, 1, 1, 4);
        var childRange2 = ws.Range(2, 1, 3, 4);
        childOfChildRange = ws.Range(3, 1, 3, 4);

        panel = new ExcelPanel(range, excelReport, templateProcessor)
        {
            Parent = globalParent,
            Children = new List<IExcelPanel>
            {
                new ExcelPanel(childRange1, excelReport, templateProcessor),
                new ExcelPanel(childRange2, excelReport, templateProcessor)
                {
                    Children = new List<IExcelPanel>
                    {
                        new ExcelPanel(childOfChildRange, excelReport, templateProcessor)
                    }
                }
            }
        };

        copiedPanel = panel.Copy(ws.Cell(5, 5));
        Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Range.LastCell());
        Assert.AreSame(globalParent, copiedPanel.Parent);

        Assert.AreEqual(2, copiedPanel.Children.Count());
        Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(5, 8), copiedPanel.Children.First().Range.LastCell());
        Assert.AreSame(copiedPanel, copiedPanel.Children.First().Parent);
        Assert.AreEqual(ws.Cell(6, 5), copiedPanel.Children.Last().Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Children.Last().Range.LastCell());
        Assert.AreSame(copiedPanel, copiedPanel.Children.Last().Parent);

        Assert.AreEqual(1, copiedPanel.Children.Last().Children.Count());
        Assert.AreEqual(ws.Cell(7, 5), copiedPanel.Children.Last().Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Children.Last().Children.First().Range.LastCell());
        Assert.AreSame(copiedPanel.Children.Last(), copiedPanel.Children.Last().Children.First().Parent);

        globalParent = new ExcelPanel(ws.Range(1, 1, 7, 7), excelReport, templateProcessor);
        range = ws.Range(1, 1, 3, 4);
        childRange1 = ws.Range(1, 1, 1, 4);
        panel = new ExcelPanel(range, excelReport, templateProcessor)
        {
            Parent = globalParent,
            Children = new List<IExcelPanel> {new ExcelPanel(childRange1, excelReport, templateProcessor)}
        };

        copiedPanel = panel.Copy(ws.Cell(5, 5));
        Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Range.LastCell());
        Assert.IsNull(copiedPanel.Parent);

        Assert.AreEqual(1, copiedPanel.Children.Count());
        Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(5, 8), copiedPanel.Children.First().Range.LastCell());
        Assert.AreSame(copiedPanel, copiedPanel.Children.First().Parent);

        globalParent = new ExcelPanel(ws.Range(1, 1, 7, 8), excelReport, templateProcessor);
        panel.Parent = globalParent;
        copiedPanel = panel.Copy(ws.Cell(5, 5), false);
        Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Range.LastCell());
        Assert.AreSame(globalParent, copiedPanel.Parent);
        Assert.AreEqual(0, copiedPanel.Children.Count());

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
        var childRange1 = ws.Range(1, 1, 2, 5);
        var childRange2 = ws.Range(3, 1, 4, 5);
        childRange2.AddToNamed("childRange2", XLScope.Worksheet);
        var namedChildRange = ws.DefinedName("childRange2");

        var childOfChildRange1 = ws.Range(2, 1, 2, 5);
        childOfChildRange1.AddToNamed("childOfChildRange1", XLScope.Worksheet);
        var childOfChildNamedRange = ws.DefinedName("childOfChildRange1");

        var childOfChildRange2 = ws.Range(4, 1, 4, 5);

        var panel = new ExcelPanel(range, excelReport, templateProcessor)
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
        Assert.AreSame(globalParent, panel.Parent);

        Assert.AreEqual(2, panel.Children.Count());
        Assert.AreEqual(ws.Cell(5, 6), panel.Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(6, 10), panel.Children.First().Range.LastCell());
        Assert.IsInstanceOf<ExcelPanel>(panel.Children.First());
        Assert.IsNotInstanceOf<INamedPanel>(panel.Children.First());
        Assert.AreSame(panel, panel.Children.First().Parent);

        Assert.AreEqual(ws.Cell(7, 6), panel.Children.Last().Range.FirstCell());
        Assert.AreEqual(ws.Cell(8, 10), panel.Children.Last().Range.LastCell());
        Assert.AreEqual("childRange2", ((INamedPanel) panel.Children.Last()).Name);
        Assert.AreSame(panel, panel.Children.First().Parent);

        Assert.AreEqual(1, panel.Children.First().Children.Count());
        Assert.AreEqual(ws.Cell(6, 6), panel.Children.First().Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(6, 10), panel.Children.First().Children.First().Range.LastCell());
        Assert.IsInstanceOf<ExcelDataSourcePanel>(panel.Children.First().Children.First());
        Assert.AreEqual("childOfChildRange1", ((INamedPanel) panel.Children.First().Children.First()).Name);
        Assert.AreSame(panel.Children.First(), panel.Children.First().Children.First().Parent);

        Assert.AreEqual(1, panel.Children.Last().Children.Count());
        Assert.AreEqual(ws.Cell(8, 6), panel.Children.Last().Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(8, 10), panel.Children.Last().Children.First().Range.LastCell());
        Assert.IsInstanceOf<ExcelPanel>(panel.Children.Last().Children.First());
        Assert.IsNotInstanceOf<INamedPanel>(panel.Children.Last().Children.First());
        Assert.AreSame(panel.Children.Last(), panel.Children.Last().Children.First().Parent);

        Assert.AreEqual(2, ws.DefinedNames.Count());

        panel.Move(ws.Cell(6, 6));

        Assert.AreEqual(ws.Cell(6, 6), panel.Range.FirstCell());
        Assert.AreEqual(ws.Cell(9, 10), panel.Range.LastCell());
        Assert.IsNull(panel.Parent);

        Assert.AreEqual(2, panel.Children.Count());
        Assert.AreEqual(ws.Cell(6, 6), panel.Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 10), panel.Children.First().Range.LastCell());
        Assert.IsInstanceOf<ExcelPanel>(panel.Children.First());
        Assert.IsNotInstanceOf<INamedPanel>(panel.Children.First());
        Assert.AreSame(panel, panel.Children.First().Parent);

        Assert.AreEqual(ws.Cell(8, 6), panel.Children.Last().Range.FirstCell());
        Assert.AreEqual(ws.Cell(9, 10), panel.Children.Last().Range.LastCell());
        Assert.AreEqual("childRange2", ((INamedPanel) panel.Children.Last()).Name);
        Assert.AreSame(panel, panel.Children.First().Parent);

        Assert.AreEqual(1, panel.Children.First().Children.Count());
        Assert.AreEqual(ws.Cell(7, 6), panel.Children.First().Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(7, 10), panel.Children.First().Children.First().Range.LastCell());
        Assert.IsInstanceOf<ExcelDataSourcePanel>(panel.Children.First().Children.First());
        Assert.AreEqual("childOfChildRange1", ((INamedPanel) panel.Children.First().Children.First()).Name);
        Assert.AreSame(panel.Children.First(), panel.Children.First().Children.First().Parent);

        Assert.AreEqual(1, panel.Children.Last().Children.Count());
        Assert.AreEqual(ws.Cell(9, 6), panel.Children.Last().Children.First().Range.FirstCell());
        Assert.AreEqual(ws.Cell(9, 10), panel.Children.Last().Children.First().Range.LastCell());
        Assert.IsInstanceOf<ExcelPanel>(panel.Children.Last().Children.First());
        Assert.IsNotInstanceOf<INamedPanel>(panel.Children.Last().Children.First());
        Assert.AreSame(panel.Children.Last(), panel.Children.Last().Children.First().Parent);

        Assert.AreEqual(2, ws.DefinedNames.Count());

        //wb.SaveAs("test.xlsx");
    }

    [Test]
    public void TestDelete()
    {
        // Deleting with moving cells up
        var wb = InitWorkBookForDeleteRangeTest();
        var ws = wb.Worksheet("Test");
        var range = ws.DefinedName("TestRange").Ranges.ElementAt(0);
        var excelReport = Substitute.For<object>();
        var templateProcessor = Substitute.For<ITemplateProcessor>();

        var panel = new ExcelPanel(range, excelReport, templateProcessor);
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

        // Deleting with moving the row up
        wb = InitWorkBookForDeleteRangeTest();
        ws = wb.Worksheet("Test");
        range = ws.DefinedName("TestRange").Ranges.ElementAt(0);

        panel = new ExcelPanel(range, excelReport, templateProcessor) {ShiftType = ShiftType.Row};
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

        // Deleting with moving cells left
        wb = InitWorkBookForDeleteRangeTest();
        ws = wb.Worksheet("Test");
        range = ws.DefinedName("TestRange").Ranges.ElementAt(0);

        panel = new ExcelPanel(range, excelReport, templateProcessor) {Type = PanelType.Horizontal};
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

        // Deleting with moving the column left
        wb = InitWorkBookForDeleteRangeTest();
        ws = wb.Worksheet("Test");
        range = ws.DefinedName("TestRange").Ranges.ElementAt(0);

        panel = new ExcelPanel(range, excelReport, templateProcessor)
            {Type = PanelType.Horizontal, ShiftType = ShiftType.Row};
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

        // Deleting without any shift
        wb = InitWorkBookForDeleteRangeTest();
        ws = wb.Worksheet("Test");
        range = ws.DefinedName("TestRange").Ranges.ElementAt(0);

        panel = new ExcelPanel(range, excelReport, templateProcessor) {ShiftType = ShiftType.NoShift};
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
        Assert.AreEqual(XLBorderStyleValues.None, range.FirstCell().Style.Border.TopBorder);
        Assert.AreEqual(XLBorderStyleValues.None, range.LastCell().Style.Border.BottomBorder);
        Assert.AreEqual(8, ws.CellsUsed(XLCellsUsedOptions.Contents).Count());
        Assert.AreEqual(belowCell1, ws.Cell(10, 6));
        Assert.AreEqual(belowCell2, ws.Cell(10, 8));
        Assert.AreEqual(rightCell1, ws.Cell(7, 8));
        Assert.AreEqual(rightCell2, ws.Cell(5, 8));
        Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
        Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
        Assert.AreEqual(leftCell1, ws.Cell(7, 4));
        Assert.AreEqual(leftCell2, ws.Cell(10, 4));

        //wb.SaveAs("test.xlsx");
    }

    [Test]
    public void TestCallReportMethod()
    {
        var report = new TestRep();
        var panel = new ExcelPanel(Substitute.For<IXLRange>(), report, Substitute.For<ITemplateProcessor>());
        var method = panel.GetType().GetMethod("CallReportMethod", BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.AreEqual($"Call {nameof(TestRep.Method1)}",
            method.Invoke(panel, new object[] {nameof(TestRep.Method1), null}));
        Assert.AreEqual($"Call {nameof(TestRep.Method1)}",
            method.Invoke(panel, new object[] {nameof(TestRep.Method1), new object[] { }}));
        ExceptionAssert.ThrowsBaseException<TargetParameterCountException>(() =>
            method.Invoke(panel, new object[] {nameof(TestRep.Method1), new object[] {"Bad param"}}));
        Assert.AreEqual($"Call {nameof(TestRep.Method2)}; params: str; 10",
            method.Invoke(panel, new object[] {nameof(TestRep.Method2), new object[] {"str", 10}}));
        Assert.AreEqual($"Call {nameof(TestRep.Method3)}; params: str; 10; str2",
            method.Invoke(panel, new object[] {nameof(TestRep.Method3), new object[] {"str", 10, "str2"}}));
        ExceptionAssert.ThrowsBaseException<TargetParameterCountException>(() =>
            method.Invoke(panel, new object[] {nameof(TestRep.Method3), new object[] {"str", 10}}));

        var eventArgs = new PanelBeforeRenderEventArgs();
        Assert.IsFalse(eventArgs.IsCanceled);
        Assert.IsNull(method.Invoke(panel, new object[] {nameof(TestRep.Method4), new object[] {eventArgs}}));
        Assert.IsTrue(eventArgs.IsCanceled);

        ExceptionAssert.ThrowsBaseException<AmbiguousMatchException>(() =>
            method.Invoke(panel, new object[] {nameof(TestRep.Method5), new object[] {eventArgs}}));
        ExceptionAssert.ThrowsBaseException<MethodNotFoundException>(
            () => method.Invoke(panel, new object[] {"Method6", null}),
            $"Cannot find public instance method \"Method6\" in type \"{report.GetType().Name}\"");
        ExceptionAssert.ThrowsBaseException<MethodNotFoundException>(
            () => method.Invoke(panel, new object[] {"BadMethod", null}),
            $"Cannot find public instance method \"BadMethod\" in type \"{report.GetType().Name}\"");

        ExceptionAssert.ThrowsBaseException<ArgumentException>(() => method.Invoke(panel, new object[] {null, null}));
        ExceptionAssert.ThrowsBaseException<ArgumentException>(() =>
            method.Invoke(panel, new object[] {string.Empty, null}));
        ExceptionAssert.ThrowsBaseException<ArgumentException>(() => method.Invoke(panel, new object[] {" ", null}));
    }

    private XLWorkbook InitWorkBookForDeleteRangeTest()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");

        var range = ws.Range(6, 5, 9, 7);
        range.AddToNamed("TestRange", XLScope.Worksheet);
        range.FirstCell().Value = "RangeStart";
        range.LastCell().Value = "RangeEnd";
        range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
        range.LastCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);

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

    private class TestRep : TestRepBase
    {
        public string Method1()
        {
            return $"Call {nameof(Method1)}";
        }

        public string Method3(string arg1, int arg2, string arg3 = null)
        {
            return $"Call {nameof(Method3)}; params: {arg1}; {arg2}; {arg3}";
        }

        public void Method4(PanelBeforeRenderEventArgs e)
        {
            e.IsCanceled = true;
        }

        public void Method5(PanelBeforeRenderEventArgs e)
        {
        }

        public void Method5()
        {
        }

        private void Method6()
        {
        }
    }

    private class TestRepBase
    {
        public string Method2(string arg1, int arg2)
        {
            return $"Call {nameof(Method2)}; params: {arg1}; {arg2}";
        }

        public void Run()
        {
            throw new NotImplementedException();
        }
    }
}