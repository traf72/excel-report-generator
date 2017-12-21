using System;

namespace ExcelReporter.Extensions
{
    internal static class StringExtensions
    {
        public static string ReplaceFirst(this string input, string oldValue, string newValue)
        {
            if (input == null)
            {
                return null;
            }

            int pos = input.IndexOf(oldValue, StringComparison.CurrentCulture);
            if (pos < 0)
            {
                return input;
            }
            return input.Substring(0, pos) + newValue + input.Substring(pos + oldValue.Length);
        }
    }
}