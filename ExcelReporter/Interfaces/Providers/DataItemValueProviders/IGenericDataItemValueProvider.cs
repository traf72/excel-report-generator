namespace ExcelReporter.Interfaces.Providers.DataItemValueProviders
{
    public interface IGenericDataItemValueProvider<in T> : IDataItemValueProvider
    {
        object GetValue(string template, T dataItem);
    }
}