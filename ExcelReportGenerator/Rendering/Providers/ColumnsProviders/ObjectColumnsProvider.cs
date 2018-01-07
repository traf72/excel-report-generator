using System;
using System.Collections.Generic;
using ExcelReportGenerator.Helpers;

namespace ExcelReportGenerator.Rendering.Providers.ColumnsProviders
{
    /// <summary>
    /// Provides columns info from any object
    /// </summary>
    internal class ObjectColumnsProvider : IColumnsProvider
    {
        private readonly IGenericColumnsProvider<Type> _typeColumnsProvider;

        public ObjectColumnsProvider(IGenericColumnsProvider<Type> typeColumnsProvider)
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