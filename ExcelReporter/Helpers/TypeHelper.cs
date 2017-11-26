using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReporter.Helpers
{
    public static class TypeHelper
    {
        public static bool IsKeyValuePair(Type type)
        {
            return type.Namespace == "System.Collections.Generic" && type.Name.StartsWith("KeyValuePair");
        }

        public static bool IsDictionaryStringObject(Type type)
        {
            if (type.Namespace == "System.Collections.Generic" && type.Name.StartsWith("IDictionary") && type.GetGenericArguments()[0] == typeof(string))
            {
                return true;
            }
            Type dictionary = type.GetInterfaces().SingleOrDefault(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IDictionary<,>));
            return dictionary != null && dictionary.GetGenericArguments()[0] == typeof(string);
        }
    }
}