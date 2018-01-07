using System.Collections.Generic;

namespace ExcelReportGenerator.Rendering.EventArgs
{
    public class DataSourceDynamicPanelEventArgs : DataSourcePanelEventArgs
    {
        public IList<ExcelDynamicColumn> Columns { get; set; }
    }
}