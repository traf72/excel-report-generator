using ExcelReporter.Exceptions;
using ExcelReporter.Interfaces.Providers;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelReporter.Implementations.Providers
{
    public class DataItemValueProvider : IDataItemValueProvider
    {
        protected HierarchicalDataItem HierarchicalDataItem { get; private set; }

        protected virtual string SelfDataItemTemplate => "di";

        public virtual object GetValue(string template, HierarchicalDataItem hierarchicalDataItem)
        {
            if (string.IsNullOrWhiteSpace(template))
            {
                throw new ArgumentException(Constants.EmptyStringParamMessage, nameof(template));
            }
            if (hierarchicalDataItem == null)
            {
                throw new ArgumentNullException(nameof(hierarchicalDataItem), Constants.NullParamMessage);
            }

            HierarchicalDataItem = hierarchicalDataItem;

            string dataItemTemplate;
            object dataItem = GetDataItemGivenHierarchy(template, out dataItemTemplate);
            if (dataItem == null)
            {
                throw new InvalidOperationException($"Data item is null for template \"{template}\"");
            }

            if (dataItemTemplate == SelfDataItemTemplate)
            {
                return dataItem;
            }

            return GetDataItemPropertyValue(dataItem, dataItemTemplate.Trim());
        }

        protected virtual object GetDataItemGivenHierarchy(string template, out string dataItemTemplate)
        {
            int lastColonIndex = template.LastIndexOf(":", StringComparison.Ordinal);
            if (lastColonIndex == -1)
            {
                dataItemTemplate = template;
                return HierarchicalDataItem.Value;
            }

            string[] parentTemplateParts = template.Substring(0, lastColonIndex).Split(':');
            HierarchicalDataItem dataItem = HierarchicalDataItem;
            foreach (string part in parentTemplateParts)
            {
                if (!part.Trim().Equals("parent", StringComparison.OrdinalIgnoreCase))
                {
                    throw new IncorrectTemplateException($"Template \"{template}\" is incorrect");
                }
                dataItem = dataItem.Parent;
            }
            dataItemTemplate = template.Substring(lastColonIndex + 1);
            return dataItem?.Value;
        }

        protected virtual object GetDataItemPropertyValue(object dataItem, string template)
        {
            Queue<string> queue = new Queue<string>(template.Split('.'));
            object obj = dataItem;
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