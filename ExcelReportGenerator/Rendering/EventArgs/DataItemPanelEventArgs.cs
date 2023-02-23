namespace ExcelReportGenerator.Rendering.EventArgs;

/// <summary>
/// Represent arguments of data item panel event
/// </summary>
public class DataItemPanelEventArgs : PanelEventArgs
{
    /// <summary>
    /// Hierarchical data item
    /// </summary>
    public HierarchicalDataItem DataItem { get; set; }
}