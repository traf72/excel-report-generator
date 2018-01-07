using ExcelReportGenerator.Helpers;
using System;
using System.Collections.Generic;

namespace ExcelReportGenerator.Rendering.Providers.ColumnsProviders
{
    /// <summary>
    /// Provides columns info from KeyValuePair
    /// </summary>
    internal class KeyValuePairColumnsProvider : IColumnsProvider
    {
        public IList<ExcelDynamicColumn> GetColumnsList(object data)
        {
            Type[] genericArgs;

            if (data == null)
            {
                genericArgs = new Type[2];
            }
            else
            {
                Type dataType = data.GetType();
                if (TypeHelper.IsKeyValuePair(dataType))
                {
                    genericArgs = dataType.GetGenericArguments();
                }
                else
                {
                    Type genericEnumerable = TypeHelper.TryGetGenericEnumerableInterface(dataType);
                    if (genericEnumerable != null)
                    {
                        genericArgs = genericEnumerable.GetGenericArguments()[0].GetGenericArguments();
                    }
                    else
                    {
                        throw new InvalidOperationException("Type of data must be KeyValuePair<TKey, TValue> or IEnumerable<KeyValuePair<TKey, TValue>>");
                    }
                }
            }

            return new List<ExcelDynamicColumn>
            {
                new ExcelDynamicColumn("Key", genericArgs[0]),
                new ExcelDynamicColumn("Value", genericArgs[1]),
            };
        }
    }
}