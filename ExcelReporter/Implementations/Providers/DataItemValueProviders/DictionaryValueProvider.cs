using ExcelReporter.Interfaces.Providers.DataItemValueProviders;
using System;
using System.Collections.Generic;

namespace ExcelReporter.Implementations.Providers.DataItemValueProviders
{
    /// <summary>
    /// Provides values from dictionary
    /// </summary>
    public class DictionaryValueProvider<TValue> : IGenericDataItemValueProvider<IDictionary<string, TValue>>
    {
        public object GetValue(string key, IDictionary<string, TValue> dataItem)
        {
            if (string.IsNullOrWhiteSpace(key))
            {
                throw new ArgumentException(Constants.EmptyStringParamMessage, nameof(key));
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