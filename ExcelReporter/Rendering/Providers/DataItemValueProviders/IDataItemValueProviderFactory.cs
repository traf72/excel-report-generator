namespace ExcelReporter.Rendering.Providers.DataItemValueProviders
{
    public interface IDataItemValueProviderFactory
    {
        IDataItemValueProvider Create(object data);
    }
}