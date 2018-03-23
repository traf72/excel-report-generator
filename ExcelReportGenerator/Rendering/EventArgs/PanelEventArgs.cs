using ClosedXML.Excel;
using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Rendering.EventArgs
{
    /// <summary>
    /// Represent arguments of panel event
    /// </summary>
    [LicenceKeyPart(L = true, R = true)]
    public class PanelEventArgs : System.EventArgs
    {
        /// <summary>
        /// Excel range
        /// </summary>
        public IXLRange Range { get; set; }
    }
}