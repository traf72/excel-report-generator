namespace ExcelReportGenerator.Rendering.Providers.DataItemValueProviders
{
    /// <summary>
    /// Provides values for data item templates
    /// </summary>
    public interface IDataItemValueProvider
    {
        /// <summary>
        /// Get value from <paramref name="dataItem"/> based on <paramref name="template"/>
        /// </summary>
        object GetValue(string template, object dataItem);
    }
}