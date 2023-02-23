namespace ExcelReportGenerator.Rendering.EventArgs;

/// <summary>
/// Represent arguments of data item panel before render event
/// </summary>
public class DataItemPanelBeforeRenderEventArgs : PanelBeforeRenderEventArgs
{
    /// <summary>
    /// Hierarchical data item
    /// </summary>
    public HierarchicalDataItem DataItem { get; set; }
}