using ClosedXML.Excel;

namespace ExcelReporter.Rendering.EventArgs
{
    public class PanelEventArgs : System.EventArgs
    {
        public IXLRange Range { get; set; }
    }
}