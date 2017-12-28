using System.Collections.Generic;

namespace ExcelReporter.Rendering.EventArgs
{
    public class DataSourceDynamicPanelEventArgs : DataSourcePanelEventArgs
    {
        public IList<ExcelDynamicColumn> Columns { get; set; }
    }
}