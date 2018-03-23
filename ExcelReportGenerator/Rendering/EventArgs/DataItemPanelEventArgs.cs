using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Rendering.EventArgs
{
    /// <summary>
    /// Represent arguments of data item panel event
    /// </summary>
    [LicenceKeyPart]
    public class DataItemPanelEventArgs : PanelEventArgs
    {
        /// <summary>
        /// Hierarchical data item
        /// </summary>
        public HierarchicalDataItem DataItem { get; set; }
    }
}