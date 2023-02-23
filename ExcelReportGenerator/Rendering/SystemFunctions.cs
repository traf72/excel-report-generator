using ExcelReportGenerator.Helpers;
using System.Collections;
using System.Globalization;

namespace ExcelReportGenerator.Rendering;

/// <summary>
/// System functions that can be called from Excel-template
/// </summary>
public class SystemFunctions
{
    /// <summary>
    /// Returns value from dictionary
    /// </summary>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="dictionary"/> or <paramref name="key"/> is null</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="dictionary"/> is not implement <see cref="IDictionary"/> interface</exception>
    /// <exception cref="KeyNotFoundException">Thrown when <paramref name="key"/> is not found in the dictionary</exception>
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

        if (dictionary is not IDictionary realDict)
        {
            throw new ArgumentException($"Parameter \"{nameof(dictionary)}\" must implement {nameof(IDictionary)} interface");
        }

        if (!realDict.Contains(key))
        {
            throw new KeyNotFoundException($"The given key \"{key}\" was not present in the dictionary");
        }

        return realDict[key];
    }

    /// <summary>
    /// Try to return value from dictionary. If key is not found or any exception occurs then returns null
    /// </summary>
    public static object TryGetDictVal(object dictionary, object key)
    {
        if (key == null || dictionary is not IDictionary realDict)
        {
            return null;
        }

        return realDict.Contains(key) ? realDict[key] : null;
    }

    /// <summary>
    /// Returns element from list by index
    /// </summary>
    /// <param name="list">List must implement <see cref="IList"/> interface</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="list"/> is null</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="list"/> is not implement <see cref="IList"/> interface</exception>
    public static object GetByIndex(object list, int index)
    {
        if (list == null)
        {
            throw new ArgumentNullException(nameof(list), ArgumentHelper.NullParamMessage);
        }
        if (list is not IList realList)
        {
            throw new ArgumentException($"Parameter \"{nameof(list)}\" must implement {nameof(IList)} interface");
        }

        return realList[index];
    }

    /// <summary>
    /// Try to return element from list by index. Returns null if any exception occurs.
    /// </summary>
    public static object TryGetByIndex(object list, int index)
    {
        if (list is not IList realList)
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

    /// <summary>
    /// Returns formatted value
    /// </summary>
    /// <param name="input">Input must implement <see cref="IFormattable"/> interface</param>
    /// <param name="formatProvider">Can implement <see cref="IFormatProvider"/>, be a <see cref="string"/>, <see cref="int"/> or null</param>
    /// <exception cref="ArgumentException">Thrown when <paramref name="input"/> is not implement <see cref="IFormattable"/> interface or <paramref name="formatProvider"/> is invalid</exception>
    public static string Format(object input, string format, object formatProvider = null)
    {
        if (input == null)
        {
            return null;
        }

        if (input is not IFormattable formattable)
        {
            throw new ArgumentException($"Parameter \"{nameof(input)}\" must implement {nameof(IFormattable)} interface");
        }

        IFormatProvider fp = formatProvider switch
        {
            null => null,
            IFormatProvider p => p,
            string str => new CultureInfo(str),
            int integer => new CultureInfo(integer),
            _ => throw new ArgumentException(
                $"Invalid type \"{formatProvider.GetType().Name}\" of {nameof(formatProvider)}")
        };

        return formattable.ToString(format, fp);
    }
}