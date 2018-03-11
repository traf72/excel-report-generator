using ClosedXML.Excel;
using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Rendering.EventArgs
{
    [LicenceKeyPart(L = true, R = true)]
    public class PanelEventArgs : System.EventArgs
    {
        public IXLRange Range { get; set; }
    }
}