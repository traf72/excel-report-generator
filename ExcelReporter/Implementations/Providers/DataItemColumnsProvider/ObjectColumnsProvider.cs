using System;
using ExcelReporter.Interfaces.Providers.DataItemColumnsProvider;
using System.Collections.Generic;

namespace ExcelReporter.Implementations.Providers.DataItemColumnsProvider
{
    internal class ObjectColumnsProvider : IDataItemColumnsProvider
    {
        private readonly IGenericDataItemColumnsProvider<Type> _typeColumnsProvider;

        public ObjectColumnsProvider(IGenericDataItemColumnsProvider<Type> typeColumnsProvider)
        {
            _typeColumnsProvider = typeColumnsProvider ?? throw new ArgumentNullException(nameof(typeColumnsProvider), Constants.NullParamMessage);
        }

        public IList<ExcelDynamicColumn> GetColumnsList(object dataItem)
        {
            return dataItem == null
                ? new List<ExcelDynamicColumn>()
                : _typeColumnsProvider.GetColumnsList(dataItem.GetType());
        }
    }
}