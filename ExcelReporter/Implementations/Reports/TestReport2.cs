using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ExcelReporter.Enums;
using ExcelReporter.Implementations.Panels;
using ExcelReporter.Implementations.Providers;
using ExcelReporter.Implementations.TemplateProcessors;
using ExcelReporter.Interfaces.Panels;
using ExcelReporter.Interfaces.Reports;
using ExcelReporter.Interfaces.TemplateProcessors;

namespace ExcelReporter.Implementations.Reports
{
    public class TestReport2 : IExcelReport
    {
        private const string TemplateName = "Template2";

        private IXLWorksheet _ws;

        public XLWorkbook Workbook { get; set; } = new XLWorkbook($@"ExcelTemplates\{TemplateName}.xlsx");

        public void Run()
        {
            //_ws = Workbook.Worksheet(1);
            //IDictionary<string, IPanel> panels = new Dictionary<string, IPanel>();
            //TemplateProcessor = new TemplateProcessor(new ReflectionParameterProvider(this));

            ////IXLNamedRange panel = _ws.NamedRange("panel");
            ////panels[panel.Name] = new ExcelDataSourcePanel("TestDataSource", "GetAllItems()", panel, report);

            ////IXLNamedRange panel = _ws.NamedRange("panel");
            ////panels[panel.Name] = new ExcelDataSourcePanel("TestDataSource", "GetAllItems()", panel, report)
            ////{
            ////    ShiftType = ShiftType.Row,
            ////};

            //IXLNamedRange panel = _ws.NamedRange("panel");
            //panels[panel.Name] = new ExcelDataSourcePanel("TestDataSource", "GetAllItems()", panel, this)
            //{
            //    ShiftType = ShiftType.NoShift,
            //};

            //foreach (KeyValuePair<string, IPanel> p in panels.OrderByDescending(p => p.Value.RenderPriority))
            //{
            //    p.Value.Render();
            //}

            //Workbook.SaveAs($@"{TemplateName}_result.xlsx");
        }

        public ITemplateProcessor TemplateProcessor { get; set; }
    }
}