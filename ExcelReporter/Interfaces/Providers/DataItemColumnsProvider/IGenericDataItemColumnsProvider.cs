using System.Collections.Generic;

namespace ExcelReporter.Interfaces.Providers.DataItemColumnsProvider
{
    internal interface IGenericDataItemColumnsProvider<in T> : IDataItemColumnsProvider
    {
        IList<ExcelDynamicColumn> GetColumnsList(T data);
    }
}