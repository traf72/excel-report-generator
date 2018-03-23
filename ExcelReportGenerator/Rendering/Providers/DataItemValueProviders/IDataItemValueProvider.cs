using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Rendering.Providers.DataItemValueProviders
{
    /// <summary>
    /// Provides values for data item templates
    /// </summary>
    [LicenceKeyPart(L = true, R = true)]
    public interface IDataItemValueProvider
    {
        /// <summary>
        /// Get value from <paramref name="dataItem"/> based on <paramref name="template"/>
        /// </summary>
        object GetValue(string template, object dataItem);
    }
}