namespace ExcelReportGenerator.Rendering.EventArgs
{
    /// <summary>
    /// Represent arguments of data source panel event
    /// </summary>
    public class DataSourcePanelEventArgs : PanelEventArgs
    {
        /// <summary>
        /// Data source panel data
        /// </summary>
        public object Data { get; set; }
    }
}