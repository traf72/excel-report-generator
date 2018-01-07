using ExcelReportGenerator.Helpers;
using System;
using System.Collections;
using System.Data;

namespace ExcelReportGenerator.Rendering.Providers.ColumnsProviders
{
    internal class ColumnsProviderFactory : IColumnsProviderFactory
    {
        public virtual IColumnsProvider Create(object data)
        {
            switch (data)
            {
                case null:
                    return null;

                case IDataReader _:
                    return new DataReaderColumnsProvider();

                case DataTable _:
                    return new DataTableColumnsProvider();

                case DataSet _:
                    return new DataSetColumnsProvider(new DataTableColumnsProvider());
            }

            Type dataType = data.GetType();
            if (TypeHelper.IsKeyValuePair(dataType))
            {
                return new KeyValuePairColumnsProvider();
            }

            Type genericEnumerable = TypeHelper.TryGetGenericEnumerableInterface(dataType);
            if (genericEnumerable != null)
            {
                Type genericType = genericEnumerable.GetGenericArguments()[0];
                if (TypeHelper.IsKeyValuePair(genericType))
                {
                    return new KeyValuePairColumnsProvider();
                }
                if (TypeHelper.IsDictionaryStringObject(genericType))
                {
                    // If data type is dictionary and type of key is String and type of value is not Object
                    Type dictionaryValueProviderRawType = typeof(DictionaryColumnsProvider<>);
                    Type dictionary = TypeHelper.TryGetGenericDictionaryInterface(genericType);
                    Type dictionaryValueProviderGenericType = dictionaryValueProviderRawType.MakeGenericType(dictionary.GetGenericArguments()[1]);
                    return (IColumnsProvider)Activator.CreateInstance(dictionaryValueProviderGenericType);
                }
                return new GenericEnumerableColumnsProvider(new TypeColumnsProvider());
            }

            if (data is IEnumerable)
            {
                return new EnumerableColumnsProvider(new TypeColumnsProvider());
            }

            return new ObjectColumnsProvider(new TypeColumnsProvider());
        }
    }
}