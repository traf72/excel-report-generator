using System.Collections.Generic;
using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Rendering.EventArgs
{
    [LicenceKeyPart(L = true)]
    public class DataSourceDynamicPanelEventArgs : DataSourcePanelEventArgs
    {
        public IList<ExcelDynamicColumn> Columns { get; set; }
    }
}