using System;

namespace ExcelReportGenerator.Converters.ExternalPropertiesConverters
{
    internal class RenderPriorityConverter : IExternalPropertyConverter<int>
    {
        public int Convert(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                throw new ArgumentException("RenderPriority property cannot be null or empty");
            }

            if (int.TryParse(input, out int renderPriority))
            {
                return renderPriority;
            }

            throw new ArgumentException($"Cannot convert value \"{input}\" to int");
        }

        object IConverter.Convert(object input)
        {
            return Convert((string)input);
        }
    }
}