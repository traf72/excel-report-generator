using System.Collections.Generic;

namespace ExcelReporter.Rendering.Providers.ColumnsProviders
{
    internal interface IGenericDataItemColumnsProvider<in T> : IDataItemColumnsProvider
    {
        IList<ExcelDynamicColumn> GetColumnsList(T data);
    }
}