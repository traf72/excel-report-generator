using ClosedXML.Excel;

namespace ExcelReportGenerator.Rendering.EventArgs
{
    /// <summary>
    /// Represent arguments of panel event
    /// </summary>
    public class PanelEventArgs : System.EventArgs
    {
        /// <summary>
        /// Excel range
        /// </summary>
        public IXLRange Range { get; set; }
    }
}