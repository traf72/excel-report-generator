using ExcelReporter.Interfaces.Providers.DataItemColumnsProvider;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReporter.Implementations.Providers.DataItemColumnsProvider
{
    internal class EnumerableColumnsProvider : IGenericDataItemColumnsProvider<IEnumerable>
    {
        private readonly IGenericDataItemColumnsProvider<Type> _typeColumnsProvider;

        public EnumerableColumnsProvider(IGenericDataItemColumnsProvider<Type> typeColumnsProvider)
        {
            _typeColumnsProvider = typeColumnsProvider ?? throw new ArgumentNullException(nameof(typeColumnsProvider), Constants.NullParamMessage);
        }

        public IList<ExcelDynamicColumn> GetColumnsList(IEnumerable data)
        {
            if (data == null)
            {
                return new List<ExcelDynamicColumn>();
            }

            object firstElem = data.Cast<object>().FirstOrDefault();
            if (firstElem == null)
            {
                return new List<ExcelDynamicColumn>();
            }

            return _typeColumnsProvider.GetColumnsList(firstElem.GetType());
        }

        IList<ExcelDynamicColumn> IDataItemColumnsProvider.GetColumnsList(object data)
        {
            return GetColumnsList((IEnumerable)data);
        }
    }
}