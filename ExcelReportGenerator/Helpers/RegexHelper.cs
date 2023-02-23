﻿using System.Text.RegularExpressions;

namespace ExcelReportGenerator.Helpers;

internal static class RegexHelper
{
    public static string SafeEscape(string value)
    {
        return value == null ? null : Regex.Escape(value);
    }
}