namespace ExcelReporter.Interfaces.Providers.DataItemColumnsProvider
{
    internal interface IDataItemColumnsProviderFactory
    {
        IDataItemColumnsProvider Create(object data);
    }
}