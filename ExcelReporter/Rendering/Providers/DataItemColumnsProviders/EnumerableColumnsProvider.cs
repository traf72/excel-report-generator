using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ExcelReporter.Helpers;

namespace ExcelReporter.Rendering.Providers.DataItemColumnsProviders
{
    /// <summary>
    /// Provides columns info from not generic enumerable
    /// </summary>
    internal class EnumerableColumnsProvider : IGenericDataItemColumnsProvider<IEnumerable>
    {
        private readonly IGenericDataItemColumnsProvider<Type> _typeColumnsProvider;

        public EnumerableColumnsProvider(IGenericDataItemColumnsProvider<Type> typeColumnsProvider)
        {
            _typeColumnsProvider = typeColumnsProvider ?? throw new ArgumentNullException(nameof(typeColumnsProvider), ArgumentHelper.NullParamMessage);
        }

        public IList<ExcelDynamicColumn> GetColumnsList(IEnumerable data)
        {
            object firstElement = data?.Cast<object>().FirstOrDefault();
            if (firstElement == null)
            {
                return new List<ExcelDynamicColumn>();
            }

            return _typeColumnsProvider.GetColumnsList(firstElement.GetType());
        }

        IList<ExcelDynamicColumn> IDataItemColumnsProvider.GetColumnsList(object data)
        {
            return GetColumnsList((IEnumerable)data);
        }
    }
}