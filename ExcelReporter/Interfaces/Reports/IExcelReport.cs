using ClosedXML.Excel;

namespace ExcelReporter.Interfaces.Reports
{
    public interface IExcelReport : IReport
    {
        XLWorkbook Workbook { get; set; }
    }
}