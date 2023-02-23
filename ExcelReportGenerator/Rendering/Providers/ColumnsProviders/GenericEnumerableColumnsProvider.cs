﻿using ExcelReportGenerator.Helpers;

namespace ExcelReportGenerator.Rendering.Providers.ColumnsProviders;

/// <summary>
/// Provides columns info from generic enumerable
/// </summary>
internal class GenericEnumerableColumnsProvider : IGenericColumnsProvider<IEnumerable<object>>
{
    private readonly IGenericColumnsProvider<Type> _typeColumnsProvider;

    public GenericEnumerableColumnsProvider(IGenericColumnsProvider<Type> typeColumnsProvider)
    {
        _typeColumnsProvider = typeColumnsProvider ?? throw new ArgumentNullException(nameof(typeColumnsProvider), ArgumentHelper.NullParamMessage);
    }

    public IList<ExcelDynamicColumn> GetColumnsList(IEnumerable<object> data)
    {
        if (data == null)
        {
            return new List<ExcelDynamicColumn>();
        }

        Type genericEnumerable = TypeHelper.TryGetGenericEnumerableInterface(data.GetType());
        return _typeColumnsProvider.GetColumnsList(genericEnumerable.GetGenericArguments()[0]);
    }

    IList<ExcelDynamicColumn> IColumnsProvider.GetColumnsList(object data)
    {
        return GetColumnsList((IEnumerable<object>)data);
    }
}