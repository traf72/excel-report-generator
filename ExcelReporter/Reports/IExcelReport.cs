using ClosedXML.Excel;

namespace ExcelReporter.Reports
{
    public interface IExcelReport : IReport
    {
        XLWorkbook Workbook { get; set; }
    }
}