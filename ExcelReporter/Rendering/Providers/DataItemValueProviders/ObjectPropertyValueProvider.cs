using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelReporter.Exceptions;
using ExcelReporter.Helpers;

namespace ExcelReporter.Rendering.Providers.DataItemValueProviders
{
    /// <summary>
    /// Provides properties values from object instance
    /// </summary>
    internal class ObjectPropertyValueProvider : IDataItemValueProvider
    {
        private string _propTemplate;
        private object _dataItem;

        protected virtual string SelfObjectTemplate => "di";

        /// <summary>
        /// Returns property value from data item object
        /// </summary>
        public virtual object GetValue(string propTemplate, object dataItem)
        {
            if (string.IsNullOrWhiteSpace(propTemplate))
            {
                throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(propTemplate));
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

                // TODO Возможно добавить также поиск публичного поля (в таком случае добавить также публичные поля при извлечении колонок)
                PropertyInfo prop = GetProperty(obj.GetType(), propName);
                obj = GetPropertyValue(prop, obj);
            }

            return obj;
        }

        /// <summary>
        /// Returns property based on type and name
        /// </summary>
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