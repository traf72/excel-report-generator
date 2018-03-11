using ExcelReportGenerator.Helpers;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Rendering
{
    [LicenceKeyPart(L = true)]
    public class SystemFunctions
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

        public static object GetByIndex(object list, int index)
        {
            if (list == null)
            {
                throw new ArgumentNullException(nameof(list), ArgumentHelper.NullParamMessage);
            }
            if (!(list is IList realList))
            {
                throw new ArgumentException($"Parameter \"{nameof(list)}\" must implement {nameof(IList)} interface");
            }

            return realList[index];
        }

        public static object TryGetByIndex(object list, int index)
        {
            if (list == null || !(list is IList realList))
            {
                return null;
            }

            try
            {
                return realList[index];
            }
            catch
            {
                return null;
            }
        }

        public static string Format(object input, string format, object formatProvider = null)
        {
            if (input == null)
            {
                return null;
            }

            if (!(input is IFormattable formattable))
            {
                throw new ArgumentException($"Parameter \"{nameof(input)}\" must implement {nameof(IFormattable)} interface");
            }

            IFormatProvider fp;
            switch (formatProvider)
            {
                case null:
                    fp = null;
                    break;

                case IFormatProvider p:
                    fp = p;
                    break;

                case string str:
                    fp = new CultureInfo(str);
                    break;

                case int integer:
                    fp = new CultureInfo(integer);
                    break;

                default:
                    throw new ArgumentException($"Invalid type \"{formatProvider.GetType().Name}\" of {nameof(formatProvider)}");
            }

            return formattable.ToString(format, fp);
        }
    }
}