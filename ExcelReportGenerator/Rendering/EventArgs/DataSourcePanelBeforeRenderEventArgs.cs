namespace ExcelReportGenerator.Rendering.EventArgs;

/// <summary>
/// Represent arguments of data source panel before render event
/// </summary>
public class DataSourcePanelBeforeRenderEventArgs : PanelBeforeRenderEventArgs
{
    /// <summary>
    /// Data source panel data
    /// </summary>
    public object Data { get; set; }
}