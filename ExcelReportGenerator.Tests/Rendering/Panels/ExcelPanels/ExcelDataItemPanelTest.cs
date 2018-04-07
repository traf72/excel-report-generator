using ClosedXML.Excel;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Rendering;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;
using System.Reflection;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels
{
    [TestClass]
    public class ExcelDataItemPanelTest
    {
        [TestMethod]
        public void TestCopy()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");
            var excelReport = Substitute.For<object>();
            var templateProcessor = Substitute.For<ITemplateProcessor>();

            IXLRange range = ws.Range(1, 1, 1, 4);

            var parentDataItem = new HierarchicalDataItem { Value = 1 };
            var dataItem = new HierarchicalDataItem { Value = "One", Parent = parentDataItem };

            var panel = new ExcelDataItemPanel(range, excelReport, templateProcessor)
            {
                RenderPriority = 10,
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.NoShift,
                BeforeRenderMethodName = "BeforeRenderMethod",
                AfterRenderMethodName = "AfterRenderMethod",
                DataItem = dataItem,
            };

            ExcelDataItemPanel copiedPanel = (ExcelDataItemPanel)panel.Copy(ws.Cell(5, 5));

            Assert.AreSame(excelReport, copiedPanel.GetType().GetField("_report", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(copiedPanel));
            Assert.AreSame(templateProcessor, copiedPanel.GetType().GetField("_templateProcessor", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(copiedPanel));
            Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Range.FirstCell());
            Assert.AreEqual(ws.Cell(5, 8), copiedPanel.Range.LastCell());
            Assert.AreEqual(10, copiedPanel.RenderPriority);
            Assert.AreEqual(PanelType.Horizontal, copiedPanel.Type);
            Assert.AreEqual(ShiftType.NoShift, copiedPanel.ShiftType);
            Assert.AreEqual("BeforeRenderMethod", copiedPanel.BeforeRenderMethodName);
            Assert.AreEqual("AfterRenderMethod", copiedPanel.AfterRenderMethodName);
            Assert.AreSame(dataItem, copiedPanel.DataItem);
            Assert.AreSame(parentDataItem, copiedPanel.DataItem.Parent);
            Assert.IsNull(copiedPanel.Parent);

            //wb.SaveAs("test.xlsx");
        }
    }
}