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
    public class TestReport : IExcelReport
    {
        private const string TemplateName = "Template";

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

            IXLNamedRange simplePanel = _ws.NamedRange("panel");
            panels[simplePanel.Name] = new DataSourcePanel("TestDataSource", "GetAllItems()", simplePanel, context);

            IXLNamedRange outer = _ws.NamedRange("Outer");
            IXLNamedRange inner = _ws.NamedRange("Inner");

            panels[outer.Name] = new DataSourcePanel("TestDataSource", "GetGroups()", outer, context);
            panels[inner.Name] = new DataSourcePanel("TestDataSource", "GetRandomItems()", inner, context);

            panels[outer.Name].Children = new List<IPanel> { panels[inner.Name] };
            panels[inner.Name].Parent = panels[outer.Name];

            panels.Remove(inner.Name);

            IXLNamedRange outer2 = _ws.NamedRange("Outer_2");
            IXLNamedRange inner2 = _ws.NamedRange("Inner_2");
            IXLNamedRange inner3 = _ws.NamedRange("Inner_3");

            panels[outer2.Name] = new DataSourcePanel("TestDataSource", "GetGroups()", outer2, context);
            panels[inner2.Name] = new DataSourcePanel("TestDataSource", "GetAllItems()", inner2, context);
            panels[inner3.Name] = new DataSourcePanel("TestDataSource", "GetRandomItems()", inner3, context);

            panels[inner2.Name].Children = new List<IPanel> { panels[inner3.Name] };
            panels[inner3.Name].Parent = panels[inner2.Name];
            panels[outer2.Name].Children = new List<IPanel> { panels[inner2.Name] };
            panels[inner2.Name].Parent = panels[outer2.Name];

            panels.Remove(inner2.Name);
            panels.Remove(inner3.Name);

            IXLNamedRange outer3 = _ws.NamedRange("Outer_3");
            IXLNamedRange inner4 = _ws.NamedRange("Inner_4");
            IXLNamedRange inner5 = _ws.NamedRange("Inner_5");

            panels[outer3.Name] = new DataSourcePanel("TestDataSource", "GetGroups()", outer3, context);
            panels[inner4.Name] = new DataSourcePanel("TestDataSource", "GetRandomItems()", inner4, context);
            panels[inner5.Name] = new DataSourcePanel("TestDataSource", "GetRandomDataItems2()", inner5, context);

            panels[outer3.Name].Children = new List<IPanel> { panels[inner4.Name], panels[inner5.Name] };
            panels[inner4.Name].Parent = panels[outer3.Name];
            panels[inner5.Name].Parent = panels[outer3.Name];

            panels.Remove(inner4.Name);
            panels.Remove(inner5.Name);

            IXLNamedRange horizPanel = _ws.NamedRange("HorizPanel");
            panels[horizPanel.Name] = new DataSourcePanel("TestDataSource", "GetAllItems()", horizPanel, context)
            {
                Type = PanelType.Horizontal
            };

            IXLNamedRange outer4 = _ws.NamedRange("Outer_4");
            IXLNamedRange inner6 = _ws.NamedRange("Inner_6");

            panels[outer4.Name] = new DataSourcePanel("TestDataSource", "GetGroups()", outer4, context);
            panels[inner6.Name] = new DataSourcePanel("TestDataSource", "GetRandomItems()", inner6, context)
            {
                Type = PanelType.Horizontal,
            };

            panels[outer4.Name].Children = new List<IPanel> { panels[inner6.Name] };
            panels[inner6.Name].Parent = panels[outer4.Name];

            panels.Remove(inner6.Name);

            foreach (KeyValuePair<string, IPanel> p in panels.OrderByDescending(p => p.Value.RenderPriority))
            {
                p.Value.Render();
            }

            Workbook.SaveAs($@"{TemplateName}_result.xlsx");
        }
    }
}