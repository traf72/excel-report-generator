namespace ExcelReportGenerator.Rendering.Providers;

/// <summary>
/// Provides values for properties templates
/// </summary>
public interface IPropertyValueProvider
{
    /// <summary>
    /// Provides property value based on <paramref name="propertyTemplate" />
    /// </summary>
    object GetValue(string propertyTemplate);
}