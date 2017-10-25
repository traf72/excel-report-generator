namespace ExcelReporter.Interfaces.Providers.DataItemValueProviders
{
    public interface IDataItemValueProviderFactory
    {
        IDataItemValueProvider Create(object data);
    }
}