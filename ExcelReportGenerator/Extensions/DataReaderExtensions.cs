﻿using System.Data;

namespace ExcelReportGenerator.Extensions;

internal static class DataReaderExtensions
{
    public static object SafeGetValue(this IDataReader reader, int columnIndex)
    {
        return !reader.IsDBNull(columnIndex) ? reader.GetValue(columnIndex) : null;
    }
}