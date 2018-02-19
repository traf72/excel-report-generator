using ClosedXML.Excel;

namespace ExcelReportGenerator.Rendering.EventArgs
{
    public class ReportRenderEventArgs : System.EventArgs
    {
        public XLWorkbook Workbook { get; set; }
    }
}