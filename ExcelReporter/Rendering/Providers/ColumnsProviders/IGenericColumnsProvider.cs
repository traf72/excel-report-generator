using System.Collections.Generic;

namespace ExcelReporter.Rendering.Providers.ColumnsProviders
{
    internal interface IGenericColumnsProvider<in T> : IColumnsProvider
    {
        IList<ExcelDynamicColumn> GetColumnsList(T data);
    }
}