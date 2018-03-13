using System;

namespace ExcelReportGenerator.Extensions
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

        public static string Reverse(this string input)
        {
            if (input == null)
            {
                return null;
            }

            int inputLength = input.Length;
            char[] inputChars = input.ToCharArray();
            char[] result = new char[inputLength];
            for (int i = 0; i < result.Length; i++)
            {
                result[i] = inputChars[inputLength - (i + 1)];
            }
            return string.Join(string.Empty, result);
        }
    }
}