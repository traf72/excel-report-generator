using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ExcelReportGenerator.Helpers;

namespace ExcelReportGenerator.Rendering.Providers.ColumnsProviders
{
    /// <summary>
    /// Provides columns info from not generic enumerable
    /// </summary>
    internal class EnumerableColumnsProvider : IGenericColumnsProvider<IEnumerable>
    {
        private readonly IGenericColumnsProvider<Type> _typeColumnsProvider;

        public EnumerableColumnsProvider(IGenericColumnsProvider<Type> typeColumnsProvider)
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

        IList<ExcelDynamicColumn> IColumnsProvider.GetColumnsList(object data)
        {
            return GetColumnsList((IEnumerable)data);
        }
    }
}