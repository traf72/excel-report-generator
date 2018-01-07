namespace ExcelReportGenerator.Rendering.Providers.DataItemValueProviders
{
    /// <summary>
    /// Provides values from data item
    /// </summary>
    public interface IGenericDataItemValueProvider<in T> : IDataItemValueProvider
    {
        /// <summary>
        /// Get value from data item based on template
        /// </summary>
        object GetValue(string template, T dataItem);
    }
}