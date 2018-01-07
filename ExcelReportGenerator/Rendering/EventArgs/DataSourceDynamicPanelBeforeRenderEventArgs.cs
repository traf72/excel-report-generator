using System.Collections.Generic;

namespace ExcelReportGenerator.Rendering.EventArgs
{
    public class DataSourceDynamicPanelBeforeRenderEventArgs : DataSourcePanelBeforeRenderEventArgs
    {
        public IList<ExcelDynamicColumn> Columns { get; set; }
    }
}