using System.Collections.Generic;

namespace ExcelReporter.Rendering.Providers.ColumnsProviders
{
    internal interface IDataItemColumnsProvider
    {
        /// <summary>
        /// Provides columns info based on data
        /// </summary>
        IList<ExcelDynamicColumn> GetColumnsList(object data);
    }
}