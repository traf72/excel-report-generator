using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelReporter.Implementations.Panels;
using ExcelReporter.Interfaces.Panels;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using ExcelReporter.Enums;

namespace ExcelReporter.Tests
{
    [TestClass]
    public class NamedPanelTest
    {
        [TestMethod]
        public void TestCopy()
        {
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLRange range = ws.Range(1, 1, 3, 4);
            range.AddToNamed("Parent", XLScope.Worksheet);
            IXLNamedRange namedRange = ws.NamedRange("Parent");

            IXLRange childRange = ws.Range(2, 1, 3, 4);
            childRange.AddToNamed("Child", XLScope.Worksheet);
            IXLNamedRange namedChildRange = ws.NamedRange("Child");

            IXLRange childOfChildRange = ws.Range(3, 1, 3, 4);
            childOfChildRange.AddToNamed("ChildOfChild", XLScope.Worksheet);
            IXLNamedRange namedChildOfChildRange = ws.NamedRange("ChildOfChild");

            var panel = new NamedPanel(namedRange, null)
            {
                Children = new List<IPanel>
                {
                    new NamedPanel(namedChildRange, null)
                    {
                        Children = new List<IPanel>
                        {
                            new DataSourcePanel("fn:DataSource:Method()", namedChildOfChildRange, null)
                        }
                    }
                }
            };

            INamedPanel copiedPanel = (INamedPanel) panel.Copy(ws.Cell(5, 5));
            Assert.IsTrue(Regex.IsMatch(copiedPanel.Name, @"Parent_[0-9a-f]{32}"));
            Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Range.FirstCell());
            Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Range.LastCell());
            Assert.IsNull(copiedPanel.Parent);

            Assert.AreEqual(1, copiedPanel.Children.Count());
            Assert.IsTrue(Regex.IsMatch(((INamedPanel)copiedPanel.Children.First()).Name, @"Parent_[0-9a-f]{32}_Child"));
            Assert.AreEqual(ws.Cell(6, 5), copiedPanel.Children.First().Range.FirstCell());
            Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Children.First().Range.LastCell());
            Assert.AreSame(copiedPanel, copiedPanel.Children.First().Parent);

            Assert.AreEqual(1, copiedPanel.Children.First().Children.Count());
            Assert.IsTrue(Regex.IsMatch(((INamedPanel)copiedPanel.Children.First().Children.First()).Name, @"Parent_[0-9a-f]{32}_Child_ChildOfChild"));
            Assert.IsInstanceOfType(copiedPanel.Children.First().Children.First(), typeof(DataSourcePanel));
            Assert.AreEqual(ws.Cell(7, 5), copiedPanel.Children.First().Children.First().Range.FirstCell());
            Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Children.First().Children.First().Range.LastCell());
            Assert.AreSame(copiedPanel.Children.First(), copiedPanel.Children.First().Children.First().Parent);

            namedRange.Delete();
            namedChildRange.Delete();
            copiedPanel.Delete();
            copiedPanel.Children.First().Delete();

            IPanel globalParent = new Panel(ws.Range(1, 1, 20, 20), null);
            range = ws.Range(1, 1, 3, 4);
            range.AddToNamed("Parent", XLScope.Worksheet);
            namedRange = ws.NamedRange("Parent");

            IXLRange childRange1 = ws.Range(1, 1, 1, 4);
            childRange1.AddToNamed("Child", XLScope.Worksheet);
            namedChildRange = ws.NamedRange("Child");

            IXLRange childRange2 = ws.Range(2, 1, 3, 4);

            childOfChildRange = ws.Range(3, 1, 3, 4);
            childOfChildRange.AddToNamed("ChildOfChild", XLScope.Worksheet);
            namedChildOfChildRange = ws.NamedRange("ChildOfChild");

            panel = new NamedPanel(namedRange, null)
            {
                Parent = globalParent,
                Children = new List<IPanel>
                {
                    new NamedPanel(namedChildRange, null),
                    new Panel(childRange2, null)
                    {
                        Children = new List<IPanel>
                        {
                            new NamedPanel(namedChildOfChildRange, null)
                        }
                    },
                },
            };

            copiedPanel = (INamedPanel) panel.Copy(ws.Cell(5, 5));
            Assert.IsTrue(Regex.IsMatch(copiedPanel.Name, @"Parent_[0-9a-f]{32}"));
            Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Range.FirstCell());
            Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Range.LastCell());
            Assert.AreSame(globalParent, copiedPanel.Parent);

            Assert.AreEqual(2, copiedPanel.Children.Count());
            Assert.IsTrue(Regex.IsMatch(((INamedPanel)copiedPanel.Children.First()).Name, @"Parent_[0-9a-f]{32}_Child"));
            Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Children.First().Range.FirstCell());
            Assert.AreEqual(ws.Cell(5, 8), copiedPanel.Children.First().Range.LastCell());
            Assert.AreSame(copiedPanel, copiedPanel.Children.First().Parent);
            Assert.IsInstanceOfType(copiedPanel.Children.Last(), typeof(Panel));
            Assert.IsNotInstanceOfType(copiedPanel.Children.Last(), typeof(NamedPanel));
            Assert.AreEqual(ws.Cell(6, 5), copiedPanel.Children.Last().Range.FirstCell());
            Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Children.Last().Range.LastCell());
            Assert.AreSame(copiedPanel, copiedPanel.Children.Last().Parent);

            Assert.AreEqual(1, copiedPanel.Children.Last().Children.Count());
            Assert.IsTrue(Regex.IsMatch(((INamedPanel)copiedPanel.Children.Last().Children.First()).Name, @"ChildOfChild_[0-9a-f]{32}"));
            Assert.AreEqual(ws.Cell(7, 5), copiedPanel.Children.Last().Children.First().Range.FirstCell());
            Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Children.Last().Children.First().Range.LastCell());
            Assert.AreSame(copiedPanel.Children.Last(), copiedPanel.Children.Last().Children.First().Parent);

            namedRange.Delete();
            namedChildRange.Delete();
            namedChildOfChildRange.Delete();
            copiedPanel.Delete();
            copiedPanel.Children.First().Delete();
            copiedPanel.Children.Last().Children.First().Delete();

            globalParent = new Panel(ws.Range(1, 1, 7, 7), null);
            range = ws.Range(1, 1, 3, 4);
            range.AddToNamed("Parent", XLScope.Worksheet);
            namedRange = ws.NamedRange("Parent");

            childRange = ws.Range(1, 1, 1, 4);
            childRange.AddToNamed("Child", XLScope.Worksheet);
            namedChildRange = ws.NamedRange("Child");

            panel = new NamedPanel(namedRange, null)
            {
                Parent = globalParent,
                Children = new List<IPanel> { new NamedPanel(namedChildRange, null) },
            };

            copiedPanel = (INamedPanel) panel.Copy(ws.Cell(5, 5));
            Assert.IsTrue(Regex.IsMatch(copiedPanel.Name, @"Parent_[0-9a-f]{32}"));
            Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Range.FirstCell());
            Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Range.LastCell());
            Assert.IsNull(copiedPanel.Parent);

            Assert.AreEqual(1, copiedPanel.Children.Count());
            Assert.IsTrue(Regex.IsMatch(((INamedPanel)copiedPanel.Children.First()).Name, @"Parent_[0-9a-f]{32}_Child"));
            Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Children.First().Range.FirstCell());
            Assert.AreEqual(ws.Cell(5, 8), copiedPanel.Children.First().Range.LastCell());
            Assert.AreSame(copiedPanel, copiedPanel.Children.First().Parent);

            copiedPanel = (INamedPanel)panel.Copy(ws.Cell(5, 5), false);
            Assert.IsTrue(Regex.IsMatch(copiedPanel.Name, @"Parent_[0-9a-f]{32}"));
            Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Range.FirstCell());
            Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Range.LastCell());
            Assert.IsNull(copiedPanel.Parent);
            Assert.AreEqual(0, copiedPanel.Children.Count());

            //wb.SaveAs(@"D:\Temp\copy.xlsx");
        }

        [TestMethod]
        public void TestCopyWithName()
        {
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLRange range = ws.Range(1, 1, 3, 4);
            range.AddToNamed("Parent", XLScope.Worksheet);
            IXLNamedRange namedRange = ws.NamedRange("Parent");

            IXLRange childRange1 = ws.Range(1, 1, 2, 4);
            childRange1.AddToNamed("Child", XLScope.Worksheet);
            IXLNamedRange namedChildRange = ws.NamedRange("Child");

            IXLRange childRange2 = ws.Range(3, 1, 3, 4);

            IXLRange childOfChildRange1 = ws.Range(1, 1, 1, 4);
            childOfChildRange1.AddToNamed("ChildOfChild1", XLScope.Worksheet);
            IXLNamedRange namedChildOfChildRange1 = ws.NamedRange("ChildOfChild1");

            IXLRange childOfChildRange2 = ws.Range(3, 1, 3, 4);
            childOfChildRange2.AddToNamed("ChildOfChild2", XLScope.Worksheet);
            IXLNamedRange namedChildOfChildRange2 = ws.NamedRange("ChildOfChild2");

            var panel = new NamedPanel(namedRange, null)
            {
                Children = new List<IPanel>
                {
                    new NamedPanel(namedChildRange, null)
                    {
                        Children = new List<IPanel>
                        {
                            new NamedPanel(namedChildOfChildRange1, null)
                        }
                    },
                    new Panel(childRange2, null)
                    {
                        Children = new List<IPanel>
                        {
                            new NamedPanel(namedChildOfChildRange2, null)
                        }
                    }
                }
            };

            INamedPanel copiedPanel = panel.Copy(ws.Cell(5, 5), "Copied");
            Assert.IsTrue(Regex.IsMatch(copiedPanel.Name, "Copied"));
            Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Range.FirstCell());
            Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Range.LastCell());
            Assert.IsNull(copiedPanel.Parent);

            Assert.AreEqual(2, copiedPanel.Children.Count());
            Assert.IsTrue(Regex.IsMatch(((INamedPanel) copiedPanel.Children.First()).Name, "Copied_Child"));
            Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Children.First().Range.FirstCell());
            Assert.AreEqual(ws.Cell(6, 8), copiedPanel.Children.First().Range.LastCell());
            Assert.AreSame(copiedPanel, copiedPanel.Children.First().Parent);
            Assert.IsInstanceOfType(copiedPanel.Children.Last(), typeof(Panel));
            Assert.IsNotInstanceOfType(copiedPanel.Children.Last(), typeof(NamedPanel));
            Assert.AreEqual(ws.Cell(7, 5), copiedPanel.Children.Last().Range.FirstCell());
            Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Children.Last().Range.LastCell());
            Assert.AreSame(copiedPanel, copiedPanel.Children.Last().Parent);

            Assert.AreEqual(1, copiedPanel.Children.First().Children.Count());
            Assert.IsTrue(Regex.IsMatch(((INamedPanel)copiedPanel.Children.First().Children.First()).Name, "Copied_Child_ChildOfChild1"));
            Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Children.First().Children.First().Range.FirstCell());
            Assert.AreEqual(ws.Cell(5, 8), copiedPanel.Children.First().Children.First().Range.LastCell());
            Assert.AreSame(copiedPanel.Children.First(), copiedPanel.Children.First().Children.First().Parent);

            Assert.AreEqual(1, copiedPanel.Children.Last().Children.Count());
            Assert.IsTrue(Regex.IsMatch(((INamedPanel)copiedPanel.Children.Last().Children.First()).Name, "ChildOfChild2_[0-9a-f]{32}"));
            Assert.AreEqual(ws.Cell(7, 5), copiedPanel.Children.Last().Children.First().Range.FirstCell());
            Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Children.Last().Children.First().Range.LastCell());
            Assert.AreSame(copiedPanel.Children.Last(), copiedPanel.Children.Last().Children.First().Parent);

            copiedPanel = panel.Copy(ws.Cell(5, 5), "Copied2", false);
            Assert.IsTrue(Regex.IsMatch(copiedPanel.Name, "Copied2"));
            Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Range.FirstCell());
            Assert.AreEqual(ws.Cell(7, 8), copiedPanel.Range.LastCell());
            Assert.IsNull(copiedPanel.Parent);
            Assert.AreEqual(0, copiedPanel.Children.Count());

            //wb.SaveAs(@"D:\Temp\copy.xlsx");
        }

        [TestMethod]
        public void TestMove()
        {
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLRange range = ws.Range(1, 1, 4, 5);
            range.AddToNamed("parentRange", XLScope.Worksheet);
            IXLNamedRange namedParentRange = ws.NamedRange("parentRange");

            IXLRange childRange1 = ws.Range(1, 1, 2, 5);
            IXLRange childRange2 = ws.Range(3, 1, 4, 5);
            childRange2.AddToNamed("childRange2", XLScope.Worksheet);
            IXLNamedRange namedChildRange = ws.NamedRange("childRange2");

            IXLRange childOfChildRange1 = ws.Range(2, 1, 2, 5);
            childOfChildRange1.AddToNamed("childOfChildRange1", XLScope.Worksheet);
            IXLNamedRange childOfChildNamedRange = ws.NamedRange("childOfChildRange1");

            IXLRange childOfChildRange2 = ws.Range(4, 1, 4, 5);

            var panel = new NamedPanel(namedParentRange, null)
            {
                Children = new List<IPanel>
                {
                    new Panel(childRange1, null)
                    {
                        Children = new List<IPanel>
                        {
                            new DataSourcePanel("fn:DataSource:Method()", childOfChildNamedRange, null)
                        }
                    },
                    new NamedPanel(namedChildRange, null)
                    {
                        Children = new List<IPanel>
                        {
                            new Panel(childOfChildRange2, null)
                        }
                    },
                }
            };

            IPanel globalParent = new Panel(ws.Range(1, 1, 8, 10), null);

            panel.Children.First().Children.First().Parent = panel.Children.First();
            panel.Children.Last().Children.First().Parent = panel.Children.Last();
            panel.Children.ForEach(c => c.Parent = panel);
            panel.Parent = globalParent;

            panel.Move(ws.Cell(5, 6));

            Assert.AreEqual(ws.Cell(5, 6), panel.Range.FirstCell());
            Assert.AreEqual(ws.Cell(8, 10), panel.Range.LastCell());
            Assert.AreEqual("parentRange", ((INamedPanel)panel).Name);
            Assert.AreSame(globalParent, panel.Parent);

            Assert.AreEqual(2, panel.Children.Count());
            Assert.AreEqual(ws.Cell(5, 6), panel.Children.First().Range.FirstCell());
            Assert.AreEqual(ws.Cell(6, 10), panel.Children.First().Range.LastCell());
            Assert.IsInstanceOfType(panel.Children.First(), typeof(Panel));
            Assert.IsNotInstanceOfType(panel.Children.First(), typeof(INamedPanel));
            Assert.AreSame(panel, panel.Children.First().Parent);

            Assert.AreEqual(ws.Cell(7, 6), panel.Children.Last().Range.FirstCell());
            Assert.AreEqual(ws.Cell(8, 10), panel.Children.Last().Range.LastCell());
            Assert.AreEqual("childRange2", ((INamedPanel)panel.Children.Last()).Name);
            Assert.AreSame(panel, panel.Children.First().Parent);

            Assert.AreEqual(1, panel.Children.First().Children.Count());
            Assert.AreEqual(ws.Cell(6, 6), panel.Children.First().Children.First().Range.FirstCell());
            Assert.AreEqual(ws.Cell(6, 10), panel.Children.First().Children.First().Range.LastCell());
            Assert.IsInstanceOfType(panel.Children.First().Children.First(), typeof(DataSourcePanel));
            Assert.AreEqual("childOfChildRange1", ((INamedPanel)panel.Children.First().Children.First()).Name);
            Assert.AreSame(panel.Children.First(), panel.Children.First().Children.First().Parent);

            Assert.AreEqual(1, panel.Children.Last().Children.Count());
            Assert.AreEqual(ws.Cell(8, 6), panel.Children.Last().Children.First().Range.FirstCell());
            Assert.AreEqual(ws.Cell(8, 10), panel.Children.Last().Children.First().Range.LastCell());
            Assert.IsInstanceOfType(panel.Children.Last().Children.First(), typeof(Panel));
            Assert.IsNotInstanceOfType(panel.Children.Last().Children.First(), typeof(INamedPanel));
            Assert.AreSame(panel.Children.Last(), panel.Children.Last().Children.First().Parent);

            Assert.AreEqual(3, ws.NamedRanges.Count());

            panel.Move(ws.Cell(6, 6));

            Assert.AreEqual(ws.Cell(6, 6), panel.Range.FirstCell());
            Assert.AreEqual(ws.Cell(9, 10), panel.Range.LastCell());
            Assert.IsNull(panel.Parent);

            Assert.AreEqual(2, panel.Children.Count());
            Assert.AreEqual(ws.Cell(6, 6), panel.Children.First().Range.FirstCell());
            Assert.AreEqual(ws.Cell(7, 10), panel.Children.First().Range.LastCell());
            Assert.IsInstanceOfType(panel.Children.First(), typeof(Panel));
            Assert.IsNotInstanceOfType(panel.Children.First(), typeof(INamedPanel));
            Assert.AreSame(panel, panel.Children.First().Parent);

            Assert.AreEqual(ws.Cell(8, 6), panel.Children.Last().Range.FirstCell());
            Assert.AreEqual(ws.Cell(9, 10), panel.Children.Last().Range.LastCell());
            Assert.AreEqual("childRange2", ((INamedPanel)panel.Children.Last()).Name);
            Assert.AreSame(panel, panel.Children.First().Parent);

            Assert.AreEqual(1, panel.Children.First().Children.Count());
            Assert.AreEqual(ws.Cell(7, 6), panel.Children.First().Children.First().Range.FirstCell());
            Assert.AreEqual(ws.Cell(7, 10), panel.Children.First().Children.First().Range.LastCell());
            Assert.IsInstanceOfType(panel.Children.First().Children.First(), typeof(DataSourcePanel));
            Assert.AreEqual("childOfChildRange1", ((INamedPanel)panel.Children.First().Children.First()).Name);
            Assert.AreSame(panel.Children.First(), panel.Children.First().Children.First().Parent);

            Assert.AreEqual(1, panel.Children.Last().Children.Count());
            Assert.AreEqual(ws.Cell(9, 6), panel.Children.Last().Children.First().Range.FirstCell());
            Assert.AreEqual(ws.Cell(9, 10), panel.Children.Last().Children.First().Range.LastCell());
            Assert.IsInstanceOfType(panel.Children.Last().Children.First(), typeof(Panel));
            Assert.IsNotInstanceOfType(panel.Children.Last().Children.First(), typeof(INamedPanel));
            Assert.AreSame(panel.Children.Last(), panel.Children.Last().Children.First().Parent);

            Assert.AreEqual(3, ws.NamedRanges.Count());
        }

        [TestMethod]
        public void TestGetNearestNamedPanelTest()
        {
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLRange range = ws.Range(1, 1, 3, 4);
            range.AddToNamed("Parent", XLScope.Worksheet);
            IXLNamedRange namedRange = ws.NamedRange("Parent");

            IXLRange childRange = ws.Range(2, 1, 3, 4);

            IXLRange childOfChildRange = ws.Range(3, 1, 3, 4);
            childOfChildRange.AddToNamed("ChildOfChild", XLScope.Worksheet);
            IXLNamedRange childOfChildNamedRange = ws.NamedRange("ChildOfChild");

            IPanel childOfChildPanel = new NamedPanel(childOfChildNamedRange, null);
            IPanel childPanel = new Panel(childRange, null);
            IPanel parentPanel = new NamedPanel(namedRange, null);

            MethodInfo method = typeof(NamedPanel).GetMethod("GetNearestNamedParent", BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.IsNull(method.Invoke(childOfChildPanel, null));

            childOfChildPanel.Parent = childPanel;
            Assert.IsNull(method.Invoke(childOfChildPanel, null));

            childPanel.Parent = parentPanel;
            Assert.AreSame(parentPanel, method.Invoke(childOfChildPanel, null));
        }

        [TestMethod]
        public void TestRemoveName()
        {
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLRange range = ws.Range(1, 1, 3, 4);
            range.AddToNamed("Parent", XLScope.Worksheet);
            IXLNamedRange namedRange = ws.NamedRange("Parent");
            INamedPanel parentPanel = new NamedPanel(namedRange, null);

            IXLRange childRange1 = ws.Range(1, 1, 1, 4);
            childRange1.AddToNamed("Child", XLScope.Worksheet);
            IXLNamedRange namedChildRange = ws.NamedRange("Child");
            INamedPanel childPanel1 = new NamedPanel(namedChildRange, null);
            childPanel1.Parent = parentPanel;

            IXLRange childRange2 = ws.Range(2, 1, 3, 4);
            IPanel childPanel2 = new Panel(childRange2, null);
            childPanel2.Parent = parentPanel;

            parentPanel.Children = new List<IPanel> {childPanel1, childPanel2};

            IXLRange childOfChild1Range = ws.Range(1, 1, 1, 4);
            childOfChild1Range.AddToNamed("ChildOfChild1", XLScope.Worksheet);
            IXLNamedRange namedChildOfChild1RangeRange = ws.NamedRange("ChildOfChild1");
            INamedPanel childOfChild1Panel = new NamedPanel(namedChildOfChild1RangeRange, null);
            childOfChild1Panel.Parent = childPanel1;
            childPanel1.Children = new List<IPanel> { childOfChild1Panel };

            IXLRange childOfChild2Range = ws.Range(3, 1, 3, 4);
            childOfChild2Range.AddToNamed("ChildOfChild2", XLScope.Worksheet);
            IXLNamedRange namedChildOfChild2RangeRange = ws.NamedRange("ChildOfChild2");
            INamedPanel childOfChild2Panel = new NamedPanel(namedChildOfChild2RangeRange, null);
            childOfChild2Panel.Parent = childPanel2;
            childPanel2.Children = new List<IPanel> { childOfChild2Panel };

            parentPanel.RemoveName();
            Assert.AreEqual(3, ws.NamedRanges.Count());
            Assert.IsNull(ws.NamedRanges.SingleOrDefault(r => r.Name == "Parent"));

            range.AddToNamed("Parent", XLScope.Worksheet);
            Assert.AreEqual(4, ws.NamedRanges.Count());
            Assert.IsNotNull(ws.NamedRanges.SingleOrDefault(r => r.Name == "Parent"));

            parentPanel.RemoveName(true);
            Assert.AreEqual(0, ws.NamedRanges.Count());
        }

        [TestMethod]
        public void TestRemoveAllNamesRecursive()
        {
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLRange range = ws.Range(1, 1, 3, 4);
            range.AddToNamed("Parent", XLScope.Worksheet);
            IXLNamedRange namedRange = ws.NamedRange("Parent");
            INamedPanel parentPanel = new NamedPanel(namedRange, null);

            IXLRange childRange1 = ws.Range(1, 1, 1, 4);
            childRange1.AddToNamed("Child", XLScope.Worksheet);
            IXLNamedRange namedChildRange = ws.NamedRange("Child");
            INamedPanel childPanel1 = new NamedPanel(namedChildRange, null);
            childPanel1.Parent = parentPanel;

            IXLRange childRange2 = ws.Range(2, 1, 3, 4);
            IPanel childPanel2 = new Panel(childRange2, null);
            childPanel2.Parent = parentPanel;

            parentPanel.Children = new List<IPanel> { childPanel1, childPanel2 };

            IXLRange childOfChild1Range = ws.Range(1, 1, 1, 4);
            childOfChild1Range.AddToNamed("ChildOfChild1", XLScope.Worksheet);
            IXLNamedRange namedChildOfChild1RangeRange = ws.NamedRange("ChildOfChild1");
            INamedPanel childOfChild1Panel = new NamedPanel(namedChildOfChild1RangeRange, null);
            childOfChild1Panel.Parent = childPanel1;
            childPanel1.Children = new List<IPanel> { childOfChild1Panel };

            IXLRange childOfChild2Range = ws.Range(3, 1, 3, 4);
            childOfChild2Range.AddToNamed("ChildOfChild2", XLScope.Worksheet);
            IXLNamedRange namedChildOfChild2RangeRange = ws.NamedRange("ChildOfChild2");
            INamedPanel childOfChild2Panel = new NamedPanel(namedChildOfChild2RangeRange, null);
            childOfChild2Panel.Parent = childPanel2;
            childPanel2.Children = new List<IPanel> { childOfChild2Panel };

            NamedPanel.RemoveAllNamesRecursive(parentPanel);
            Assert.AreEqual(0, ws.NamedRanges.Count());
        }

        [TestMethod]
        public void TestDelete()
        {
            // Удаление со сдвигом ячеек вверх
            XLWorkbook wb = InitWorkBookForDeleteRangeTest();
            IXLWorksheet ws = wb.Worksheet("Test");
            IXLNamedRange parentRange = ws.NamedRange("Parent");
            IXLNamedRange childRange = ws.NamedRange("Child");
            Assert.AreEqual(2, ws.NamedRanges.Count());

            var panel = new NamedPanel(parentRange, null)
            {
                Children = new List<IPanel> { new NamedPanel(childRange, null) }
            };
            panel.Delete();

            IXLCell rangeStartCell = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RangeStart");
            IXLCell rangeEndCell = ws.Cells().SingleOrDefault(c => c.Value.ToString() == "RangeEnd");
            IXLCell belowCell1 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_1");
            IXLCell belowCell2 = ws.Cells().Single(c => c.Value.ToString() == "BelowCell_2");
            IXLCell rightCell1 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_1");
            IXLCell rightCell2 = ws.Cells().Single(c => c.Value.ToString() == "RightCell_2");
            IXLCell aboveCell1 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_1");
            IXLCell aboveCell2 = ws.Cells().Single(c => c.Value.ToString() == "AboveCell_2");
            IXLCell leftCell1 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_1");
            IXLCell leftCell2 = ws.Cells().Single(c => c.Value.ToString() == "LeftCell_2");

            Assert.IsNull(rangeStartCell);
            Assert.IsNull(rangeEndCell);
            Assert.AreEqual(8, ws.CellsUsed().Count());
            Assert.AreEqual(belowCell1, ws.Cell(6, 6));
            Assert.AreEqual(belowCell2, ws.Cell(10, 8));
            Assert.AreEqual(rightCell1, ws.Cell(7, 8));
            Assert.AreEqual(rightCell2, ws.Cell(5, 8));
            Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
            Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
            Assert.AreEqual(leftCell1, ws.Cell(7, 4));
            Assert.AreEqual(leftCell2, ws.Cell(10, 4));
            Assert.AreEqual(0, ws.NamedRanges.Count());

            // Удаление со сдвигом строки вверх
            wb = InitWorkBookForDeleteRangeTest();
            ws = wb.Worksheet("Test");
            parentRange = ws.NamedRange("Parent");
            childRange = ws.NamedRange("Child");
            Assert.AreEqual(2, ws.NamedRanges.Count());

            panel = new NamedPanel(parentRange, null)
            {
                Children = new List<IPanel> { new NamedPanel(childRange, null) },
                ShiftType = ShiftType.Row,
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
            Assert.AreEqual(6, ws.CellsUsed().Count());
            Assert.AreEqual(belowCell1, ws.Cell(6, 6));
            Assert.AreEqual(belowCell2, ws.Cell(6, 8));
            Assert.AreEqual(rightCell2, ws.Cell(5, 8));
            Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
            Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
            Assert.AreEqual(leftCell2, ws.Cell(6, 4));
            Assert.AreEqual(0, ws.NamedRanges.Count());

            // Удаление со сдвигом ячеек влево
            wb = InitWorkBookForDeleteRangeTest();
            ws = wb.Worksheet("Test");
            parentRange = ws.NamedRange("Parent");
            childRange = ws.NamedRange("Child");
            Assert.AreEqual(2, ws.NamedRanges.Count());

            panel = new NamedPanel(parentRange, null)
            {
                Children = new List<IPanel> { new NamedPanel(childRange, null) },
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
            Assert.AreEqual(8, ws.CellsUsed().Count());
            Assert.AreEqual(belowCell1, ws.Cell(10, 6));
            Assert.AreEqual(belowCell2, ws.Cell(10, 8));
            Assert.AreEqual(rightCell1, ws.Cell(7, 5));
            Assert.AreEqual(rightCell2, ws.Cell(5, 8));
            Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
            Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
            Assert.AreEqual(leftCell1, ws.Cell(7, 4));
            Assert.AreEqual(leftCell2, ws.Cell(10, 4));
            Assert.AreEqual(0, ws.NamedRanges.Count());

            // Удаление со сдвигом колонки влево
            wb = InitWorkBookForDeleteRangeTest();
            ws = wb.Worksheet("Test");
            parentRange = ws.NamedRange("Parent");
            childRange = ws.NamedRange("Child");
            Assert.AreEqual(2, ws.NamedRanges.Count());

            panel = new NamedPanel(parentRange, null)
            {
                Children = new List<IPanel> { new NamedPanel(childRange, null) },
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.Row,
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
            Assert.AreEqual(6, ws.CellsUsed().Count());
            Assert.AreEqual(belowCell2, ws.Cell(10, 5));
            Assert.AreEqual(rightCell1, ws.Cell(7, 5));
            Assert.AreEqual(rightCell2, ws.Cell(5, 5));
            Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
            Assert.AreEqual(leftCell1, ws.Cell(7, 4));
            Assert.AreEqual(leftCell2, ws.Cell(10, 4));
            Assert.AreEqual(0, ws.NamedRanges.Count());

            // Удаление без сдвига
            wb = InitWorkBookForDeleteRangeTest();
            ws = wb.Worksheet("Test");
            parentRange = ws.NamedRange("Parent");
            childRange = ws.NamedRange("Child");
            Assert.AreEqual(2, ws.NamedRanges.Count());

            panel = new NamedPanel(parentRange, null)
            {
                Children = new List<IPanel> { new NamedPanel(childRange, null) },
                ShiftType = ShiftType.NoShift,
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
            Assert.AreEqual(8, ws.CellsUsed().Count());
            Assert.AreEqual(belowCell1, ws.Cell(10, 6));
            Assert.AreEqual(belowCell2, ws.Cell(10, 8));
            Assert.AreEqual(rightCell1, ws.Cell(7, 8));
            Assert.AreEqual(rightCell2, ws.Cell(5, 8));
            Assert.AreEqual(aboveCell1, ws.Cell(5, 6));
            Assert.AreEqual(aboveCell2, ws.Cell(5, 4));
            Assert.AreEqual(leftCell1, ws.Cell(7, 4));
            Assert.AreEqual(leftCell2, ws.Cell(10, 4));
            Assert.AreEqual(0, ws.NamedRanges.Count());

            //wb.SaveAs(@"D:\Temp\copy.xlsx");
        }

        private XLWorkbook InitWorkBookForDeleteRangeTest()
        {
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");

            IXLRange range = ws.Range(6, 5, 9, 7);
            range.AddToNamed("Parent", XLScope.Worksheet);
            range.FirstCell().Value = "RangeStart";
            range.LastCell().Value = "RangeEnd";
            range.FirstCell().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            range.LastCell().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);

            IXLRange childRange = ws.Range(7, 5, 8, 7);
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
}