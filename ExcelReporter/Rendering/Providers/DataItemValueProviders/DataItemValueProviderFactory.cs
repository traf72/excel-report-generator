using ExcelReporter.Helpers;
using System;
using System.Collections.Generic;
using System.Data;

namespace ExcelReporter.Rendering.Providers.DataItemValueProviders
{
    internal class DataItemValueProviderFactory : IDataItemValueProviderFactory
    {
        public virtual IDataItemValueProvider Create(object data)
        {
            switch (data)
            {
                case null:
                    return new ObjectPropertyValueProvider();

                case DataRow _:
                    return new DataRowValueProvider();

                case IDataReader _:
                    return new DataReaderValueProvider();

                case IDictionary<string, object> _:
                    return new DictionaryValueProvider<object>();
            }

            Type dataType = data.GetType();
            if (TypeHelper.IsDictionaryStringObject(dataType))
            {
                // If data type is dictionary and type of key is String and type of value is not Object
                Type dictionaryValueProviderRawType = typeof(DictionaryValueProvider<>);
                Type dictionary = TypeHelper.TryGetGenericDictionaryInterface(dataType);
                Type dictionaryValueProviderGenericType = dictionaryValueProviderRawType.MakeGenericType(dictionary.GetGenericArguments()[1]);
                return (IDataItemValueProvider)Activator.CreateInstance(dictionaryValueProviderGenericType);
            }

            return new ObjectPropertyValueProvider();
        }
    }
}