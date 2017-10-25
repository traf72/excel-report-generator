using ExcelReporter.Exceptions;
using ExcelReporter.Interfaces.Providers;
using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelReporter.Interfaces.Providers.DataItemValueProviders;

namespace ExcelReporter.Implementations.Providers.DataItemValueProviders
{
    public class ObjectPropertyValueProvider : IDataItemValueProvider
    {
        private string _propTemplate;
        private object _dataItem;

        protected virtual string SelfObjectTemplate => "di";

        public virtual object GetValue(string propTemplate, object dataItem)
        {
            if (string.IsNullOrWhiteSpace(propTemplate))
            {
                throw new ArgumentException(Constants.EmptyStringParamMessage, nameof(propTemplate));
            }

            _propTemplate = propTemplate.Trim();
            _dataItem = dataItem;

            if (_propTemplate == SelfObjectTemplate)
            {
                return _dataItem;
            }

            return GetValue();
        }

        private object GetValue()
        {
            Queue<string> queue = new Queue<string>(_propTemplate.Split('.'));
            object obj = _dataItem;
            while (queue.Count > 0)
            {
                string propName = queue.Dequeue();
                if (obj == null)
                {
                    throw new InvalidOperationException($"Cannot get property \"{propName}\" because object is null");
                }

                PropertyInfo prop = GetProperty(obj.GetType(), propName);
                obj = GetPropertyValue(prop, obj);
            }

            return obj;
        }

        protected virtual PropertyInfo GetProperty(Type type, string name)
        {
            PropertyInfo prop = type.GetProperty(name, BindingFlags.Instance | BindingFlags.Public);
            if (prop == null)
            {
                throw new MemberNotFoundException($"Cannot find public instance property \"{name}\" in class \"{type.Name}\" and all its parents");
            }
            return prop;
        }

        protected virtual object GetPropertyValue(PropertyInfo prop, object obj)
        {
            return prop.GetValue(obj);
        }
    }
}