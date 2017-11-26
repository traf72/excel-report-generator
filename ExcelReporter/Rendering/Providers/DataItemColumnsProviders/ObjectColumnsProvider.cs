using System;
using System.Collections.Generic;
using ExcelReporter.Helpers;

namespace ExcelReporter.Rendering.Providers.DataItemColumnsProviders
{
    /// <summary>
    /// Provides columns info from any object
    /// </summary>
    internal class ObjectColumnsProvider : IDataItemColumnsProvider
    {
        private readonly IGenericDataItemColumnsProvider<Type> _typeColumnsProvider;

        public ObjectColumnsProvider(IGenericDataItemColumnsProvider<Type> typeColumnsProvider)
        {
            _typeColumnsProvider = typeColumnsProvider ?? throw new ArgumentNullException(nameof(typeColumnsProvider), ArgumentHelper.NullParamMessage);
        }

        public IList<ExcelDynamicColumn> GetColumnsList(object dataItem)
        {
            return dataItem == null
                ? new List<ExcelDynamicColumn>()
                : _typeColumnsProvider.GetColumnsList(dataItem.GetType());
        }
    }
}