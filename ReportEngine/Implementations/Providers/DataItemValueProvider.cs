using ReportEngine.Exceptions;
using ReportEngine.Interfaces.Providers;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace ReportEngine.Implementations.Providers
{
    public class DataItemValueProvider : IDataItemValueProvider
    {
        public object GetValue(string template, object dataItem)
        {
            if (string.IsNullOrWhiteSpace(template))
            {
                throw new ArgumentException(Constants.EmptyStringParamMessage, nameof(template));
            }

            if (template == "di") //dataitem
            {
                return dataItem;
            }

            Queue<string> queue = new Queue<string>(template.Split('.'));
            object obj = dataItem;
            while (queue.Count > 0)
            {
                string propName = queue.Dequeue();
                if (obj == null)
                {
                    throw new InvalidOperationException($"Cannot get property \"{propName}\" because object is null");
                }
                Type objType = obj.GetType();
                PropertyInfo prop = objType.GetProperty(propName);
                if (prop == null)
                {
                    throw new MemberNotFoundException($"Cannot find property with name \"{propName}\" in class \"{objType.Name}\"");
                }
                obj = prop.GetValue(obj);
            }

            return obj;
        }
    }
}