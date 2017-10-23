using ClosedXML.Excel;
using ExcelReporter.Implementations.Panels.Excel;
using ExcelReporter.Interfaces.Panels;
using ExcelReporter.Interfaces.Panels.Excel;
using ExcelReporter.Interfaces.Reports;
using ExcelReporter.Interfaces.TemplateProcessors;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReporter.Implementations.Reports
{
    public class TestReport4 : IExcelReport
    {
        private const string TemplateName = "Template4";

        private IXLWorksheet _ws;

        public XLWorkbook Workbook { get; set; } = new XLWorkbook($@"ExcelTemplates\{TemplateName}.xlsx");

        public void Run()
        {
            _ws = Workbook.Worksheet(1);
            IDictionary<string, IExcelPanel> panels = new Dictionary<string, IExcelPanel>();
            //DefaultTemplateProcessor = new DefaultTemplateProcessor(new ReflectionParameterProvider(this), new TypeInstanceProvider(null), new DataItemValueProvider());

            IXLNamedRange simplePanel = _ws.NamedRange("panel");
            panels[simplePanel.Name] = new ExcelDataSourcePanel("fn:TestDataSource:GetAllItems()", simplePanel, this);

            //IXLNamedRange outer = _ws.NamedRange("Outer");
            //IXLNamedRange inner = _ws.NamedRange("Inner");

            //panels[outer.Name] = new ExcelDataSourcePanel("TestDataSource", "GetGroups()", outer, this);
            //panels[inner.Name] = new ExcelDataSourcePanel("TestDataSource", "GetRandomItems()", inner, this);

            //panels[outer.Name].Children = new List<IPanel> { panels[inner.Name] };
            //panels[inner.Name].Parent = panels[outer.Name];

            //panels.Remove(inner.Name);

            //IXLNamedRange outer2 = _ws.NamedRange("Outer_2");
            //IXLNamedRange inner2 = _ws.NamedRange("Inner_2");
            //IXLNamedRange inner3 = _ws.NamedRange("Inner_3");

            //panels[outer2.Name] = new ExcelDataSourcePanel("TestDataSource", "GetGroups()", outer2, this);
            //panels[inner2.Name] = new ExcelDataSourcePanel("TestDataSource", "GetAllItems()", inner2, this);
            //panels[inner3.Name] = new ExcelDataSourcePanel("TestDataSource", "GetRandomItems()", inner3, this);

            //panels[inner2.Name].Children = new List<IPanel> { panels[inner3.Name] };
            //panels[inner3.Name].Parent = panels[inner2.Name];
            //panels[outer2.Name].Children = new List<IPanel> { panels[inner2.Name] };
            //panels[inner2.Name].Parent = panels[outer2.Name];

            //panels.Remove(inner2.Name);
            //panels.Remove(inner3.Name);

            //IXLNamedRange outer3 = _ws.NamedRange("Outer_3");
            //IXLNamedRange inner4 = _ws.NamedRange("Inner_4");
            //IXLNamedRange inner5 = _ws.NamedRange("Inner_5");

            //panels[outer3.Name] = new ExcelDataSourcePanel("TestDataSource", "GetGroups()", outer3, this);
            //panels[inner4.Name] = new ExcelDataSourcePanel("TestDataSource", "GetRandomItems()", inner4, this);
            //panels[inner5.Name] = new ExcelDataSourcePanel("TestDataSource", "GetRandomDataItems2()", inner5, this);

            //panels[outer3.Name].Children = new List<IPanel> { panels[inner4.Name], panels[inner5.Name] };
            //panels[inner4.Name].Parent = panels[outer3.Name];
            //panels[inner5.Name].Parent = panels[outer3.Name];

            //panels.Remove(inner4.Name);
            //panels.Remove(inner5.Name);

            //IXLNamedRange horizPanel = _ws.NamedRange("HorizPanel");
            //panels[horizPanel.Name] = new ExcelDataSourcePanel("TestDataSource", "GetAllItems()", horizPanel, this)
            //{
            //    Type = PanelType.Horizontal
            //};

            //IXLNamedRange outer4 = _ws.NamedRange("Outer_4");
            //IXLNamedRange inner6 = _ws.NamedRange("Inner_6");

            //panels[outer4.Name] = new ExcelDataSourcePanel("TestDataSource", "GetGroups()", outer4, this);
            //panels[inner6.Name] = new ExcelDataSourcePanel("TestDataSource", "GetRandomItems()", inner6, this)
            //{
            //    Type = PanelType.Horizontal,
            //};

            //panels[outer4.Name].Children = new List<IPanel> { panels[inner6.Name] };
            //panels[inner6.Name].Parent = panels[outer4.Name];

            //panels.Remove(inner6.Name);

            foreach (KeyValuePair<string, IExcelPanel> p in panels.OrderByDescending(p => p.Value.RenderPriority))
            {
                p.Value.Render();
            }

            Workbook.SaveAs($@"{TemplateName}_result.xlsx");
        }

        public ITemplateProcessor TemplateProcessor { get; set; }
    }
}