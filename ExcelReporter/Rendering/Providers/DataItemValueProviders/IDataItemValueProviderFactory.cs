namespace ExcelReporter.Rendering.Providers.DataItemValueProviders
{
    internal interface IDataItemValueProviderFactory
    {
        IDataItemValueProvider Create(object data);
    }
}