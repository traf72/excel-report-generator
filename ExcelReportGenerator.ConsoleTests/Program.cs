using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Rendering;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Rendering.Providers;
using ExcelReportGenerator.Tests.Rendering;
using ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests;

namespace ExcelReportGenerator.ConsoleTests
{
    class Program
    {
        static void Main(string[] args)
        {
            //XLWorkbook wb = new XLWorkbook();
            //var ws = wb.AddWorksheet("Test");
            //var range = ws.Range(1, 1, 1, 2);
            //range.AddToNamed("TestRange");
            //IXLNamedRange namedRange = wb.NamedRange("TestRange");
            //namedRange.Comment = $"Line1{Environment.NewLine}Line2";
            //wb.SaveAs("test.xlsx");

            //var input = new XLWorkbook(@"d:\temp\TestSheet.xlsx");
            //var generator = new DefaultReportGenerator(new object());
            //generator.Render(input, null);

            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(2, 2, 2, 6);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(2, 2).Value = "{di:Name}";
            ws.Cell(2, 3).Value = "{di:Date}";
            ws.Cell(2, 4).Value = "{di:Sum}";
            ws.Cell(2, 5).Value = "{di:Contacts.Phone}";
            ws.Cell(2, 6).Value = "{di:Contacts.Fax}";

            const int dataCount = 3000;
            //IList<TestItem> data = new List<TestItem>(dataCount);
            //for (int i = 0; i < dataCount; i++)
            //{
            //    data.Add(new TestItem($"Name_{i}", DateTime.Now.AddHours(1), i + 10, new Contacts($"Phone_{i}", $"Fax_{i}")));
            //}

            var data = new DataTable();
            data.Columns.Add("Name", typeof(string));
            data.Columns.Add("Date", typeof(DateTime));
            data.Columns.Add("Sum", typeof(decimal));
            data.Columns.Add("Contacts.Phone", typeof(string));
            data.Columns.Add("Contacts.Fax", typeof(string));

            for (int i = 0; i < dataCount; i++)
            {
                data.Rows.Add($"Name_{i}", DateTime.Now.AddHours(1), i + 10, $"Phone_{i}", $"Fax_{i}");
            }

            var panel = new ExcelDataSourcePanel(data, ws.NamedRange("TestRange"), report, report.TemplateProcessor)
            {
                //ShiftType = ShiftType.NoShift,
                ShiftType = ShiftType.Row,
            };

            Stopwatch sw = Stopwatch.StartNew();

            //var ws2 = range.Worksheet;
            //var row = range.Worksheet.Row(999);
            //for (int i = 0; i < dataCount; i++)
            //{
            //    //range.InsertRowsBelow(1, true);
            //    //range.Worksheet.Row(range.FirstRow().RowNumber()).InsertRowsAbove(range.RowCount());
            //    //range.Worksheet.Row(range.LastRow().RowNumber()).InsertRowsBelow(range.RowCount());
            //    //ws2.Row(1).InsertRowsBelow(1);
            //    //range.Worksheet.Row(range.FirstRow().RowNumber()).InsertRowsAbove(1);
            //    //range.Worksheet.Row(range.FirstRow().RowNumber()).InsertRowsAbove(1);
            //    //range.Worksheet.Row(1000).InsertRowsAbove(range.RowCount());

            //    range.FirstCell().WorksheetRow().InsertRowsAbove(1);
            //    //range.Worksheet.Row(i + 1).InsertRowsAbove(range.RowCount());

            //    //row.RowBelow().InsertRowsAbove(range.RowCount());
            //    //range.Worksheet.Row(1).
            //    //range.Worksheet.Row(2000).InsertRowsBelow(1);
            //}

            //foreach (var item in data)
            //{
            //    var a = 10;
            //}

            //range.Worksheet.Row(1).InsertRowsAbove(3000);

            panel.Render();

            sw.Stop();

            //Assert.AreEqual(ws.Range(2, 2, 4, 2), resultRange);

            report.Workbook.SaveAs("test.xlsx");
        }
    }
}
