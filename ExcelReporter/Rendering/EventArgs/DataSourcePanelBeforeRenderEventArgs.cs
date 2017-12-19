namespace ExcelReporter.Rendering.EventArgs
{
    public class DataSourcePanelBeforeRenderEventArgs : PanelBeforeRenderEventArgs
    {
        public object Data { get; set; }
    }
}