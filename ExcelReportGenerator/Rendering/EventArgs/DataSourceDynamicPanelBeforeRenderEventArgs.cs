using System.Collections.Generic;

namespace ExcelReportGenerator.Rendering.EventArgs
{
    /// <summary>
    /// Represent arguments of data source dynamic panel before render event
    /// </summary>
    public class DataSourceDynamicPanelBeforeRenderEventArgs : DataSourcePanelBeforeRenderEventArgs
    {
        /// <summary>
        /// Data source dynamic panel columns
        /// </summary>
        public IList<ExcelDynamicColumn> Columns { get; set; }
    }
}