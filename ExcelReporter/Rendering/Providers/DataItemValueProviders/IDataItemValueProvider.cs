namespace ExcelReporter.Rendering.Providers.DataItemValueProviders
{
    /// <summary>
    /// Provides values from data item
    /// </summary>
    public interface IDataItemValueProvider
    {
        /// <summary>
        /// Get value from data item based on template
        /// </summary>
        object GetValue(string template, object dataItem);
    }
}