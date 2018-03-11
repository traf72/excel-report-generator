using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Rendering.EventArgs
{
    [LicenceKeyPart]
    public class DataItemPanelEventArgs : PanelEventArgs
    {
        public HierarchicalDataItem DataItem { get; set; }
    }
}