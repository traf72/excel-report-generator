using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Rendering.Providers
{
    [LicenceKeyPart(U = true)]
    public interface IPropertyValueProvider
    {
        /// <summary>
        /// Provides property value base on string property template
        /// </summary>
        object GetValue(string propertyTemplate);
    }
}