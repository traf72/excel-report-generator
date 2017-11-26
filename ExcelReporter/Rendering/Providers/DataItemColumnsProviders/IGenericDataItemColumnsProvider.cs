using System.Collections.Generic;

namespace ExcelReporter.Rendering.Providers.DataItemColumnsProviders
{
    internal interface IGenericDataItemColumnsProvider<in T> : IDataItemColumnsProvider
    {
        IList<ExcelDynamicColumn> GetColumnsList(T data);
    }
}