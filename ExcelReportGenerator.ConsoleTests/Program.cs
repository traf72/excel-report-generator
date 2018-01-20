using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using ExcelReportGenerator.Rendering;
using ExcelReportGenerator.Rendering.Providers;

namespace ExcelReportGenerator.ConsoleTests
{
    class Program
    {
        static void Main(string[] args)
        {
            XLWorkbook wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Test");
            var range = ws.Range(1, 1, 1, 2);
            range.AddToNamed("TestRange");
            IXLNamedRange namedRange = wb.NamedRange("TestRange");
            namedRange.Comment = $"Line1{Environment.NewLine}Line2";
            wb.SaveAs("test.xlsx");

            //var input = new XLWorkbook(@"d:\temp\TestSheet.xlsx");
            //var generator = new DefaultReportGenerator(new object());
            //generator.Render(input, null);
        }
    }
}
