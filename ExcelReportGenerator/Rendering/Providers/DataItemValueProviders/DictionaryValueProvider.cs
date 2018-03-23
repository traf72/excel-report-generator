using System;
using System.Collections.Generic;
using ExcelReportGenerator.Helpers;

namespace ExcelReportGenerator.Rendering.Providers.DataItemValueProviders
{
    // Provides values from dictionary
    internal class DictionaryValueProvider<TValue> : IGenericDataItemValueProvider<IDictionary<string, TValue>>
    {
        public object GetValue(string key, IDictionary<string, TValue> dataItem)
        {
            if (string.IsNullOrWhiteSpace(key))
            {
                throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(key));
            }

            if (dataItem.TryGetValue(key, out TValue value))
            {
                return value;
            }
            throw new KeyNotFoundException($"Key \"{key}\" was not found in dictionary");
        }

        public object GetValue(string key, object dataItem)
        {
            return GetValue(key, (IDictionary<string, TValue>)dataItem);
        }
    }
}