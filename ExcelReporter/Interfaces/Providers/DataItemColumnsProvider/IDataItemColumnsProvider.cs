using System.Collections.Generic;

namespace ExcelReporter.Interfaces.Providers.DataItemColumnsProvider
{
    internal interface IDataItemColumnsProvider
    {
        /// <summary>
        /// Provides columns info based on data
        /// </summary>
        IList<ExcelDynamicColumn> GetColumnsList(object data);
    }
}