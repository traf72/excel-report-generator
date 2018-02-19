using ClosedXML.Excel;

namespace ExcelReportGenerator.Rendering.EventArgs
{
    public class WorksheetRenderEventArgs : System.EventArgs
    {
        public IXLWorksheet Worksheet { get; set; }
    }
}