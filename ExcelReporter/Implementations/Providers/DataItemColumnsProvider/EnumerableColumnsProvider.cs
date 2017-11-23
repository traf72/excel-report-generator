using ExcelReporter.Interfaces.Providers.DataItemColumnsProvider;
using System.Collections.Generic;

namespace ExcelReporter.Implementations.Providers.DataItemColumnsProvider
{
    internal class EnumerableColumnsProvider : IGenericDataItemColumnsProvider<IEnumerable<object>>
    {
        private readonly IDataItemColumnsProvider _typeColumnsProvider;

        public EnumerableColumnsProvider(IDataItemColumnsProvider typeColumnsProvider)
        {
            _typeColumnsProvider = typeColumnsProvider;
        }

        public IList<ExcelDynamicColumn> GetColumnsList(IEnumerable<object> data)
        {
            return data == null
                ? new List<ExcelDynamicColumn>()
                : _typeColumnsProvider.GetColumnsList(data.GetType().GetGenericArguments()[0]);
        }

        IList<ExcelDynamicColumn> IDataItemColumnsProvider.GetColumnsList(object data)
        {
            return GetColumnsList((IEnumerable<object>)data);
        }
    }
}