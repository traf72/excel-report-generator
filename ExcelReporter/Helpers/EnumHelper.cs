using System;

namespace ExcelReporter.Helpers
{
    public static class EnumHelper
    {
        public static TEnum Parse<TEnum>(string value, bool ignoreCase = true) where TEnum : struct, IConvertible
        {
            return (TEnum)Enum.Parse(typeof(TEnum), value, ignoreCase);
        }
    }
}