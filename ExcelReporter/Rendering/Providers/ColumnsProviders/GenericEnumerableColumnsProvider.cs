using System;
using System.Collections.Generic;
using ExcelReporter.Helpers;

namespace ExcelReporter.Rendering.Providers.ColumnsProviders
{
    /// <summary>
    /// Provides columns info from generic enumerable
    /// </summary>
    internal class GenericEnumerableColumnsProvider : IGenericDataItemColumnsProvider<IEnumerable<object>>
    {
        private readonly IGenericDataItemColumnsProvider<Type> _typeColumnsProvider;

        public GenericEnumerableColumnsProvider(IGenericDataItemColumnsProvider<Type> typeColumnsProvider)
        {
            _typeColumnsProvider = typeColumnsProvider ?? throw new ArgumentNullException(nameof(typeColumnsProvider), ArgumentHelper.NullParamMessage);
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