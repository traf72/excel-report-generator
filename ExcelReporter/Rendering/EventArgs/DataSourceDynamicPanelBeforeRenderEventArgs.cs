using System.Collections.Generic;

namespace ExcelReporter.Rendering.EventArgs
{
    public class DataSourceDynamicPanelBeforeRenderEventArgs : DataSourcePanelBeforeRenderEventArgs
    {
        public IList<ExcelDynamicColumn> Columns { get; set; }
    }
}