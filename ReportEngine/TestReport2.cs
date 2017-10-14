using ClosedXML.Excel;
using ReportEngine.Enums;
using ReportEngine.Implementations.Panels;
using ReportEngine.Implementations.Providers;
using ReportEngine.Implementations.Reports;
using ReportEngine.Implementations.TemplateProcessors;
using ReportEngine.Interfaces.Panels;
using ReportEngine.Interfaces.Reports;
using System.Collections.Generic;
using System.Linq;

namespace ReportEngine
{
    public class TestReport2 : IExcelReport
    {
        private const string TemplateName = "Template2";

        private IXLWorksheet _ws;

        public XLWorkbook Workbook { get; set; } = new XLWorkbook($@"ExcelTemplates\{TemplateName}.xlsx");

        public void Run()
        {
            _ws = Workbook.Worksheet(1);
            IDictionary<string, IPanel> panels = new Dictionary<string, IPanel>();
            var context = new ExcelReportContext
            {
                //TemplateProcessor = new TemplateProcessor(new ReflectionParameterProvider(this), null, this),
                Report = this,
            };

            //IXLNamedRange panel = _ws.NamedRange("panel");
            //panels[panel.Name] = new DataSourcePanel("TestDataSource", "GetAllItems()", panel, report);

            //IXLNamedRange panel = _ws.NamedRange("panel");
            //panels[panel.Name] = new DataSourcePanel("TestDataSource", "GetAllItems()", panel, report)
            //{
            //    ShiftType = ShiftType.Row,
            //};

            IXLNamedRange panel = _ws.NamedRange("panel");
            panels[panel.Name] = new DataSourcePanel("TestDataSource", "GetAllItems()", panel, context)
            {
                ShiftType = ShiftType.NoShift,
            };

            foreach (KeyValuePair<string, IPanel> p in panels.OrderByDescending(p => p.Value.RenderPriority))
            {
                p.Value.Render();
            }

            Workbook.SaveAs($@"{TemplateName}_result.xlsx");
        }
    }
}