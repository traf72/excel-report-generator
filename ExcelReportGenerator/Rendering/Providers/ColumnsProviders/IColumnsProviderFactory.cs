namespace ExcelReportGenerator.Rendering.Providers.ColumnsProviders
{
    internal interface IColumnsProviderFactory
    {
        // Create appropriate column provider based on data 
        IColumnsProvider Create(object data);
    }
}