﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using ExcelReporter.Helpers;

namespace ExcelReporter.Rendering.Providers.DataItemValueProviders
{
    internal class DefaultDataItemValueProviderFactory : IDataItemValueProviderFactory
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
                Type dictionary = dataType.GetInterfaces().SingleOrDefault(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IDictionary<,>));
                Type dictionaryValueProviderGenericType = dictionaryValueProviderRawType.MakeGenericType(dictionary.GetGenericArguments()[1]);
                return (IDataItemValueProvider)Activator.CreateInstance(dictionaryValueProviderGenericType);
            }

            return new ObjectPropertyValueProvider();
        }
    }
}