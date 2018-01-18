using ExcelReportGenerator.Helpers;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelReportGenerator.Rendering
{
    internal static class SystemFunctions
    {
        public static object GetDictVal(object dictionary, object key)
        {
            if (dictionary == null)
            {
                throw new ArgumentNullException(nameof(dictionary), ArgumentHelper.NullParamMessage);
            }
            if (key == null)
            {
                throw new ArgumentNullException(nameof(key), ArgumentHelper.NullParamMessage);
            }

            if (!(dictionary is IDictionary realDict))
            {
                throw new ArgumentException($"Parameter \"{nameof(dictionary)}\" must implement {nameof(IDictionary)} interface");
            }

            if (!realDict.Contains(key))
            {
                throw new KeyNotFoundException($"The given key \"{key}\" was not present in the dictionary");
            }

            return realDict[key];
        }

        public static object TryGetDictVal(object dictionary, object key)
        {
            if (key == null || !(dictionary is IDictionary realDict))
            {
                return null;
            }

            return realDict.Contains(key) ? realDict[key] : null;
        }

        public static string Format(object input, string format, IFormatProvider formatProvider = null)
        {
            if (input == null)
            {
                return null;
            }

            if (!(input is IFormattable formattable))
            {
                throw new ArgumentException($"Parameter \"{nameof(input)}\" must implement {nameof(IFormattable)} interface");
            }

            return formattable.ToString(format, formatProvider);
        }
    }
}