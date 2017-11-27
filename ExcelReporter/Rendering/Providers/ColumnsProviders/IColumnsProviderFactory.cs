namespace ExcelReporter.Rendering.Providers.ColumnsProviders
{
    internal interface IDataItemColumnsProviderFactory
    {
        /// <summary>
        /// Create appropriate column provider based on data 
        /// </summary>
        IDataItemColumnsProvider Create(object data);
    }
}