using System;
using System.Reflection;

namespace ExcelReportGenerator.Extensions
{
    internal static class ParameterInfoExtensions
    {
        public static bool IsParams(this ParameterInfo parameter)
        {
            return parameter.IsDefined(typeof(ParamArrayAttribute), false);
        }

        // Support .NET 4.0
        public static bool HasDefaultValue(this ParameterInfo parameter)
        {
            return parameter.DefaultValue != DBNull.Value;
        }
    }
}