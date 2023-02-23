namespace ExcelReportGenerator.Rendering.EventArgs;

/// <summary>
/// Represent arguments of panel before render event
/// </summary>
public class PanelBeforeRenderEventArgs : PanelEventArgs
{
    /// <summary>
    /// Is rendering canceled
    /// </summary>
    public bool IsCanceled { get; set; }
}