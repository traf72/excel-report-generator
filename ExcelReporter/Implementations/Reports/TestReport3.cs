using ClosedXML.Excel;
using ExcelReporter.Enums;
using ExcelReporter.Implementations.Panels.Excel;
using ExcelReporter.Implementations.Providers;
using ExcelReporter.Implementations.Providers.DataItemValueProviders;
using ExcelReporter.Implementations.TemplateProcessors;
using ExcelReporter.Interfaces.Panels.Excel;
using ExcelReporter.Interfaces.Reports;
using ExcelReporter.Interfaces.TemplateProcessors;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReporter.Implementations.Reports
{
    public class TestReport3 : IExcelReport
    {
        private const string TemplateName = "Template3";

        private IXLWorksheet _ws;

        public XLWorkbook Workbook { get; set; } = new XLWorkbook($@"ExcelTemplates\{TemplateName}.xlsx");

        public void Run()
        {
            _ws = Workbook.Worksheet(1);
            IDictionary<string, IExcelPanel> panels = new Dictionary<string, IExcelPanel>();
            TemplateProcessor = new DefaultTemplateProcessor(new ReflectionParameterProvider(this), new MethodCallValueProvider(new TypeProvider(), this), new HierarchicalDataItemValueProvider());

            //IXLNamedRange panel = _ws.NamedRange("panel");
            //panels[panel.Name] = new ExcelDataSourcePanel("m:TestDataSource:GetAllItems()", panel, this)
            //{
            //    Type = PanelType.Horizontal
            //};

            //IXLNamedRange panel = _ws.NamedRange("panel");
            //panels[panel.Name] = new ExcelDataSourcePanel("m:TestDataSource:GetAllItems()", panel, this)
            //{
            //    Type = PanelType.Horizontal,
            //    ShiftType = ShiftType.Row,
            //};

            IXLNamedRange panel = _ws.NamedRange("panel");
            panels[panel.Name] = new ExcelDataSourcePanel("m:TestDataSource:GetAllItems()", panel, this)
            {
                Type = PanelType.Horizontal,
                ShiftType = ShiftType.NoShift,
            };

            foreach (KeyValuePair<string, IExcelPanel> p in panels.OrderByDescending(p => p.Value.RenderPriority))
            {
                p.Value.Render();
            }

            Workbook.SaveAs($@"{TemplateName}_result.xlsx");
        }

        public ITemplateProcessor TemplateProcessor { get; set; }
    }
}