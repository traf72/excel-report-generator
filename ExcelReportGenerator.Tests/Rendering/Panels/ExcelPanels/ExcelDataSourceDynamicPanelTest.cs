﻿using ClosedXML.Excel;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;
using System.Reflection;

namespace ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels
{
    [TestClass]
    public class ExcelDataSourceDynamicPanelTest
    {
        [TestMethod]
        public void TestCopyIfDataSourceTemplateIsSet()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");
            var excelReport = Substitute.For<object>();
            var templateProcessor = Substitute.For<ITemplateProcessor>();

            IXLRange range = ws.Range(1, 1, 2, 4);
            range.AddToNamed("DataPanel", XLScope.Worksheet);
            IXLNamedRange namedRange = ws.NamedRange("DataPanel");

            var panel = new ExcelDataSourceDynamicPanel("m:GetData()", namedRange, excelReport, templateProcessor)
            {
                RenderPriority = 10,
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.NoShift,
                BeforeRenderMethodName = "BeforeRenderMethod",
                AfterRenderMethodName = "AfterRenderMethod",
                BeforeDataItemRenderMethodName = "BeforeDataItemRenderMethodName",
                AfterDataItemRenderMethodName = "AfterDataItemRenderMethodName",
                GroupBy = "2,4",
                AfterDataRenderMethodName = "AfterDataRenderMethodName",
                AfterDataTemplatesRenderMethodName = "AfterDataTemplatesRenderMethodName",
                AfterHeadersRenderMethodName = "AfterHeadersRenderMethodName",
                AfterNumbersRenderMethodName = "AfterNumbersRenderMethodName",
                AfterTotalsRenderMethodName = "AfterTotalsRenderMethodName",
                AfterTotalsTemplatesRenderMethodName = "AfterTotalsTemplatesRenderMethodName",
                BeforeDataRenderMethodName = "BeforeDataRenderMethodName",
                BeforeDataTemplatesRenderMethodName = "BeforeDataTemplatesRenderMethodName",
                BeforeHeadersRenderMethodName = "BeforeHeadersRenderMethodName",
                BeforeNumbersRenderMethodName = "BeforeNumbersRenderMethodName",
                BeforeTotalsRenderMethodName = "BeforeTotalsRenderMethodName",
                BeforeTotalsTemplatesRenderMethodName = "BeforeTotalsTemplatesRenderMethodName",
            };

            ExcelDataSourceDynamicPanel copiedPanel = (ExcelDataSourceDynamicPanel)panel.Copy(ws.Cell(5, 5));

            Assert.AreSame(excelReport, copiedPanel.GetType().GetField("_report", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(copiedPanel));
            Assert.AreSame(templateProcessor, copiedPanel.GetType().GetField("_templateProcessor", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(copiedPanel));
            Assert.IsNull(copiedPanel.GetType().GetField("_data", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(copiedPanel));
            Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Range.FirstCell());
            Assert.AreEqual(ws.Cell(6, 8), copiedPanel.Range.LastCell());
            Assert.AreEqual(10, copiedPanel.RenderPriority);
            Assert.AreEqual(PanelType.Horizontal, copiedPanel.Type);
            Assert.AreEqual(ShiftType.NoShift, copiedPanel.ShiftType);
            Assert.AreEqual("BeforeRenderMethod", copiedPanel.BeforeRenderMethodName);
            Assert.AreEqual("AfterRenderMethod", copiedPanel.AfterRenderMethodName);
            Assert.AreEqual("BeforeDataItemRenderMethodName", copiedPanel.BeforeDataItemRenderMethodName);
            Assert.AreEqual("AfterDataItemRenderMethodName", copiedPanel.AfterDataItemRenderMethodName);
            Assert.AreEqual("2,4", copiedPanel.GroupBy);
            Assert.AreEqual("AfterDataRenderMethodName", copiedPanel.AfterDataRenderMethodName);
            Assert.AreEqual("AfterDataTemplatesRenderMethodName", copiedPanel.AfterDataTemplatesRenderMethodName);
            Assert.AreEqual("AfterHeadersRenderMethodName", copiedPanel.AfterHeadersRenderMethodName);
            Assert.AreEqual("AfterNumbersRenderMethodName", copiedPanel.AfterNumbersRenderMethodName);
            Assert.AreEqual("AfterTotalsRenderMethodName", copiedPanel.AfterTotalsRenderMethodName);
            Assert.AreEqual("AfterTotalsTemplatesRenderMethodName", copiedPanel.AfterTotalsTemplatesRenderMethodName);
            Assert.AreEqual("BeforeDataRenderMethodName", copiedPanel.BeforeDataRenderMethodName);
            Assert.AreEqual("BeforeDataTemplatesRenderMethodName", copiedPanel.BeforeDataTemplatesRenderMethodName);
            Assert.AreEqual("BeforeHeadersRenderMethodName", copiedPanel.BeforeHeadersRenderMethodName);
            Assert.AreEqual("BeforeNumbersRenderMethodName", copiedPanel.BeforeNumbersRenderMethodName);
            Assert.AreEqual("BeforeTotalsRenderMethodName", copiedPanel.BeforeTotalsRenderMethodName);
            Assert.AreEqual("BeforeTotalsTemplatesRenderMethodName", copiedPanel.BeforeTotalsTemplatesRenderMethodName);

            Assert.IsNull(copiedPanel.Parent);

            //wb.SaveAs("test.xlsx");
        }

        [TestMethod]
        public void TestCopyIfDataIsSet()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Test");
            var excelReport = Substitute.For<object>();
            var templateProcessor = Substitute.For<ITemplateProcessor>();

            IXLRange range = ws.Range(1, 1, 2, 4);
            range.AddToNamed("DataPanel", XLScope.Worksheet);
            IXLNamedRange namedRange = ws.NamedRange("DataPanel");

            object[] data = { 1, "One" };
            var panel = new ExcelDataSourceDynamicPanel(data, namedRange, excelReport, templateProcessor)
            {
                RenderPriority = 10,
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.NoShift,
                BeforeRenderMethodName = "BeforeRenderMethod",
                AfterRenderMethodName = "AfterRenderMethod",
                BeforeDataItemRenderMethodName = "BeforeDataItemRenderMethodName",
                AfterDataItemRenderMethodName = "AfterDataItemRenderMethodName",
                GroupBy = "2,4",
                AfterDataRenderMethodName = "AfterDataRenderMethodName",
                AfterDataTemplatesRenderMethodName = "AfterDataTemplatesRenderMethodName",
                AfterHeadersRenderMethodName = "AfterHeadersRenderMethodName",
                AfterNumbersRenderMethodName = "AfterNumbersRenderMethodName",
                AfterTotalsRenderMethodName = "AfterTotalsRenderMethodName",
                AfterTotalsTemplatesRenderMethodName = "AfterTotalsTemplatesRenderMethodName",
                BeforeDataRenderMethodName = "BeforeDataRenderMethodName",
                BeforeDataTemplatesRenderMethodName = "BeforeDataTemplatesRenderMethodName",
                BeforeHeadersRenderMethodName = "BeforeHeadersRenderMethodName",
                BeforeNumbersRenderMethodName = "BeforeNumbersRenderMethodName",
                BeforeTotalsRenderMethodName = "BeforeTotalsRenderMethodName",
                BeforeTotalsTemplatesRenderMethodName = "BeforeTotalsTemplatesRenderMethodName",
            };

            ExcelDataSourceDynamicPanel copiedPanel = (ExcelDataSourceDynamicPanel)panel.Copy(ws.Cell(5, 5));

            Assert.AreSame(excelReport, copiedPanel.GetType().GetField("_report", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(copiedPanel));
            Assert.AreSame(templateProcessor, copiedPanel.GetType().GetField("_templateProcessor", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(copiedPanel));
            Assert.IsNull(copiedPanel.GetType().GetField("_dataSourceTemplate", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(copiedPanel));
            Assert.AreSame(data, copiedPanel.GetType().GetField("_data", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(copiedPanel));
            Assert.AreEqual(ws.Cell(5, 5), copiedPanel.Range.FirstCell());
            Assert.AreEqual(ws.Cell(6, 8), copiedPanel.Range.LastCell());
            Assert.AreEqual(10, copiedPanel.RenderPriority);
            Assert.AreEqual(PanelType.Horizontal, copiedPanel.Type);
            Assert.AreEqual(ShiftType.NoShift, copiedPanel.ShiftType);
            Assert.AreEqual("BeforeRenderMethod", copiedPanel.BeforeRenderMethodName);
            Assert.AreEqual("AfterRenderMethod", copiedPanel.AfterRenderMethodName);
            Assert.AreEqual("BeforeDataItemRenderMethodName", copiedPanel.BeforeDataItemRenderMethodName);
            Assert.AreEqual("AfterDataItemRenderMethodName", copiedPanel.AfterDataItemRenderMethodName);
            Assert.AreEqual("2,4", copiedPanel.GroupBy);
            Assert.AreEqual("AfterDataRenderMethodName", copiedPanel.AfterDataRenderMethodName);
            Assert.AreEqual("AfterDataTemplatesRenderMethodName", copiedPanel.AfterDataTemplatesRenderMethodName);
            Assert.AreEqual("AfterHeadersRenderMethodName", copiedPanel.AfterHeadersRenderMethodName);
            Assert.AreEqual("AfterNumbersRenderMethodName", copiedPanel.AfterNumbersRenderMethodName);
            Assert.AreEqual("AfterTotalsRenderMethodName", copiedPanel.AfterTotalsRenderMethodName);
            Assert.AreEqual("AfterTotalsTemplatesRenderMethodName", copiedPanel.AfterTotalsTemplatesRenderMethodName);
            Assert.AreEqual("BeforeDataRenderMethodName", copiedPanel.BeforeDataRenderMethodName);
            Assert.AreEqual("BeforeDataTemplatesRenderMethodName", copiedPanel.BeforeDataTemplatesRenderMethodName);
            Assert.AreEqual("BeforeHeadersRenderMethodName", copiedPanel.BeforeHeadersRenderMethodName);
            Assert.AreEqual("BeforeNumbersRenderMethodName", copiedPanel.BeforeNumbersRenderMethodName);
            Assert.AreEqual("BeforeTotalsRenderMethodName", copiedPanel.BeforeTotalsRenderMethodName);
            Assert.AreEqual("BeforeTotalsTemplatesRenderMethodName", copiedPanel.BeforeTotalsTemplatesRenderMethodName);
            Assert.IsNull(copiedPanel.Parent);

            //wb.SaveAs("test.xlsx");
        }
    }
}