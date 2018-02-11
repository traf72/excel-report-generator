using System;
using System.Linq;

namespace ExcelReportGenerator.Extensions
{
    internal static class TypeExtensions
    {
        private static readonly Type[] _numericTypes = {
            typeof(byte), typeof(ushort), typeof(uint), typeof(ulong), typeof(sbyte), typeof(short), typeof(int), typeof(long), typeof(float), typeof(double), typeof(decimal),
            typeof(byte?), typeof(ushort?), typeof(uint?), typeof(ulong?), typeof(sbyte?), typeof(short?), typeof(int?), typeof(long?), typeof(float?), typeof(double?), typeof(decimal?)
        };

        private static readonly Type[] _extendedPrimitiveTypes;

        static TypeExtensions()
        {
            _extendedPrimitiveTypes = _numericTypes.Concat(new[]
            {
                typeof(char), typeof(bool), typeof(string), typeof(Guid), typeof(DateTime),
                typeof(char?), typeof(bool?), typeof(Guid?), typeof(DateTime?)
            }).ToArray();
        }

        public static bool IsNumeric(this Type type)
        {
            return _numericTypes.Contains(type);
        }

        public static bool IsExtendedPrimitive(this Type type)
        {
            return _extendedPrimitiveTypes.Contains(type);
        }
    }
}