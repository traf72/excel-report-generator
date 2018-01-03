namespace ExcelReporter.Converters.ExternalPropertiesConverters
{
    internal interface IExternalPropertyConverter<out TOut> : IGenericConverter<string, TOut>
    {
    }
}