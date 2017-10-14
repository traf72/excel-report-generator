using ExcelReporter.Exceptions;
using ExcelReporter.Interfaces.Providers;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelReporter.Implementations.Providers
{
    //public class DataItemValueProviderNew : IDataItemValueProvider
    //{
    //    private readonly HierarchicalDataItem _dataItem;

    //    public DataItemValueProviderNew(HierarchicalDataItem dataItem)
    //    {
    //        if (dataItem == null)
    //        {
    //            throw new ArgumentNullException(nameof(dataItem), Constants.NullParamMessage);
    //        }

    //        _dataItem = dataItem;
    //    }

    //    public object GetValue(string template)
    //    {
    //        if (string.IsNullOrWhiteSpace(template))
    //        {
    //            throw new ArgumentException(Constants.EmptyStringParamMessage, nameof(template));
    //        }

    //        // TODO Также нужно проверять свойства в родителе
    //        if (template == "di") //dataitem
    //        {
    //            return _dataItem.DataItem;
    //        }

    //        Queue<string> queue = new Queue<string>(template.Split('.'));
    //        object obj = _dataItem;
    //        while (queue.Count > 0)
    //        {
    //            string propName = queue.Dequeue();
    //            if (obj == null)
    //            {
    //                throw new InvalidOperationException($"Cannot get property \"{propName}\" because object is null");
    //            }
    //            Type objType = obj.GetType();
    //            PropertyInfo prop = objType.GetProperty(propName);
    //            if (prop == null)
    //            {
    //                throw new MemberNotFoundException($"Cannot find property with name \"{propName}\" in class \"{objType.Name}\"");
    //            }
    //            obj = prop.GetValue(obj);
    //        }

    //        return obj;
    //    }
    //}
}