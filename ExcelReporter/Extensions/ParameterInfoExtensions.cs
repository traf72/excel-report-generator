using System;
using System.Reflection;

namespace ExcelReporter.Extensions
{
    internal static class ParameterInfoExtensions
    {
        public static bool IsParams(this ParameterInfo parameter)
        {
            return parameter.IsDefined(typeof(ParamArrayAttribute), false);
        }
    }
}