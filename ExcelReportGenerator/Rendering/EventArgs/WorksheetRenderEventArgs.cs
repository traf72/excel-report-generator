using ClosedXML.Excel;

namespace ExcelReportGenerator.Rendering.EventArgs
{
    /// <summary>
    /// Represent arguments of worksheet render event
    /// </summary>
    public class WorksheetRenderEventArgs : System.EventArgs
    {
        /// <summary>
        /// Excel worksheet
        /// </summary>
        public IXLWorksheet Worksheet { get; set; }
    }
}