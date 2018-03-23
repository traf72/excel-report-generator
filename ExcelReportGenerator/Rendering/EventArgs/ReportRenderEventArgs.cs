using ClosedXML.Excel;

namespace ExcelReportGenerator.Rendering.EventArgs
{
    /// <summary>
    /// Represent arguments of report render event
    /// </summary>
    public class ReportRenderEventArgs : System.EventArgs
    {
        /// <summary>
        /// Excel workbook
        /// </summary>
        public XLWorkbook Workbook { get; set; }
    }
}