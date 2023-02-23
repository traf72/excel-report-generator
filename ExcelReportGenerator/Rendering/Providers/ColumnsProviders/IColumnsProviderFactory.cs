namespace ExcelReportGenerator.Rendering.Providers.ColumnsProviders;

internal interface IColumnsProviderFactory
{
    /// <summary>
    /// Create appropriate column provider based on data
    /// </summary>
    IColumnsProvider Create(object data);
}