using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReportGenerator.Helpers
{
    internal static class TypeHelper
    {
        public static bool IsKeyValuePair(Type type)
        {
            return type != null && type.Namespace == "System.Collections.Generic" && type.Name.StartsWith("KeyValuePair");
        }

        public static bool IsEnumerableOfKeyValuePair(Type type)
        {
            Type genericEnumerable = TryGetGenericEnumerableInterface(type);
            return genericEnumerable != null && IsKeyValuePair(genericEnumerable.GetGenericArguments()[0]);
        }

        // Check if the type is dictionary with string keys and values of any type
        public static bool IsDictionaryStringObject(Type type)
        {
            if (type == null)
            {
                return false;
            }

            if (type.Namespace == "System.Collections.Generic" && type.Name.StartsWith("IDictionary") && type.GetGenericArguments()[0] == typeof(string))
            {
                return true;
            }
            Type dictionary = type.GetInterfaces().SingleOrDefault(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IDictionary<,>));
            return dictionary != null && dictionary.GetGenericArguments()[0] == typeof(string);
        }

        // If the type implements IEnumerable<T> returns it otherwise returns null
        public static Type TryGetGenericEnumerableInterface(Type type)
        {
            if (type == null)
            {
                return null;
            }

            if (type.IsInterface && type.Namespace == "System.Collections.Generic" && type.Name.StartsWith("IEnumerable"))
            {
                return type;
            }
            return type.GetInterfaces().SingleOrDefault(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IEnumerable<>));
        }

        // If the type implements IDictionary<TKey, TValue> returns it otherwise returns null
        public static Type TryGetGenericDictionaryInterface(Type type)
        {
            if (type == null)
            {
                return null;
            }

            if (type.IsInterface && type.Namespace == "System.Collections.Generic" && type.Name.StartsWith("IDictionary"))
            {
                return type;
            }
            return type.GetInterfaces().SingleOrDefault(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IDictionary<,>));
        }

        public static bool IsGenericEnumerable(Type type)
        {
            return TryGetGenericEnumerableInterface(type) != null;
        }

        public static Type TryGetGenericCollectionInterface(Type type)
        {
            if (type == null)
            {
                return null;
            }

            if (type.IsInterface && type.Namespace == "System.Collections.Generic" && type.Name.StartsWith("ICollection"))
            {
                return type;
            }
            return type.GetInterfaces().SingleOrDefault(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(ICollection<>));
        }
    }
}