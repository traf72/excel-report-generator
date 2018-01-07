using ClosedXML.Excel;

namespace ExcelReportGenerator.Rendering.EventArgs
{
    public class PanelEventArgs : System.EventArgs
    {
        public IXLRange Range { get; set; }
    }
}