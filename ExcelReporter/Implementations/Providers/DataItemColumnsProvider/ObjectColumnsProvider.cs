using ExcelReporter.Interfaces.Providers.DataItemColumnsProvider;
using System.Collections.Generic;

namespace ExcelReporter.Implementations.Providers.DataItemColumnsProvider
{
    internal class ObjectColumnsProvider : IDataItemColumnsProvider
    {
        private readonly IDataItemColumnsProvider _typeColumnsProvider;

        public ObjectColumnsProvider(IDataItemColumnsProvider typeColumnsProvider)
        {
            _typeColumnsProvider = typeColumnsProvider;
        }

        public IList<ExcelDynamicColumn> GetColumnsList(object dataItem)
        {
            return dataItem == null
                ? new List<ExcelDynamicColumn>()
                : _typeColumnsProvider.GetColumnsList(dataItem.GetType());
        }
    }
}