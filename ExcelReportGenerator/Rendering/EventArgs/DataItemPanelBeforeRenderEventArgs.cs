namespace ExcelReportGenerator.Rendering.EventArgs
{
    public class DataItemPanelBeforeRenderEventArgs : PanelBeforeRenderEventArgs
    {
        public HierarchicalDataItem DataItem { get; set; }
    }
}