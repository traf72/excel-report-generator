using ClosedXML.Excel;

namespace ReportEngine.Interfaces.Reports
{
    public interface IExcelReport : IReport
    {
        XLWorkbook Workbook { get; set; }
    }
}