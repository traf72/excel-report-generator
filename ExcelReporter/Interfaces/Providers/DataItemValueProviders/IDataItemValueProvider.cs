namespace ExcelReporter.Interfaces.Providers.DataItemValueProviders
{
    public interface IDataItemValueProvider
    {
        object GetValue(string template, object dataItem);
    }
}