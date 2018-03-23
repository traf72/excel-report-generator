using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Rendering.Providers
{
    /// <summary>
    /// Provides values for properties templates
    /// </summary>
    [LicenceKeyPart(U = true)]
    public interface IPropertyValueProvider
    {
        /// <summary>
        /// Provides property value based on <paramref name="propertyTemplate" />
        /// </summary>
        object GetValue(string propertyTemplate);
    }
}