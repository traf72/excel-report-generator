using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Helpers;
using System;

namespace ExcelReportGenerator.Converters.ExternalPropertiesConverters
{
    internal class PanelTypeConverter : IExternalPropertyConverter<PanelType>
    {
        public PanelType Convert(string panelType)
        {
            if (string.IsNullOrWhiteSpace(panelType))
            {
                throw new ArgumentException("PanelType property cannot be null or empty");
            }

            try
            {
                return EnumHelper.Parse<PanelType>(panelType.Trim());
            }
            catch (ArgumentException e)
            {
                throw new ArgumentException($"Value \"{panelType}\" is invalid for PanelType property", e);
            }
        }

        object IConverter.Convert(object input)
        {
            return Convert((string)input);
        }
    }
}