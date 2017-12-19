namespace ExcelReporter.Rendering.EventArgs
{
    public class PanelBeforeRenderEventArgs : PanelEventArgs
    {
        public bool IsCanceled { get; set; }
    }
}