using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Helpers;
using System;

namespace ExcelReportGenerator.Converters.ExternalPropertiesConverters
{
    internal class ShiftTypeConverter : IExternalPropertyConverter<ShiftType>
    {
        public ShiftType Convert(string shiftType)
        {
            if (string.IsNullOrWhiteSpace(shiftType))
            {
                throw new ArgumentException("ShiftType property cannot be null or empty");
            }

            try
            {
                return EnumHelper.Parse<ShiftType>(shiftType.Trim());
            }
            catch (ArgumentException e)
            {
                throw new ArgumentException($"Value \"{shiftType}\" is invalid for ShiftType property", e);
            }
        }

        object IConverter.Convert(object input)
        {
            return Convert((string)input);
        }
    }
}