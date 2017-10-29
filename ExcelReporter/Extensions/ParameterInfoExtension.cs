using System;
using System.Reflection;

namespace ExcelReporter.Extensions
{
    internal static class ParameterInfoExtension
    {
        public static bool IsParams(this ParameterInfo parameter)
        {
            return parameter.IsDefined(typeof(ParamArrayAttribute), false);
        }
    }
}