using System.Collections.Generic;
using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Rendering.EventArgs
{
    /// <summary>
    /// Represent arguments of data source dynamic panel event
    /// </summary>
    [LicenceKeyPart(L = true)]
    public class DataSourceDynamicPanelEventArgs : DataSourcePanelEventArgs
    {
        /// <summary>
        /// Data source dynamic panel columns
        /// </summary>
        public IList<ExcelDynamicColumn> Columns { get; set; }
    }
}