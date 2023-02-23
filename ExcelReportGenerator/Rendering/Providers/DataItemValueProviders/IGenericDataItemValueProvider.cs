namespace ExcelReportGenerator.Rendering.Providers.DataItemValueProviders;

/// See <see cref="IDataItemValueProvider"/>
/// <typeparam name="T">Type of data item</typeparam>
public interface IGenericDataItemValueProvider<in T> : IDataItemValueProvider
{
    /// See <see cref="IDataItemValueProvider.GetValue(string, object)"/>
    object GetValue(string template, T dataItem);
}