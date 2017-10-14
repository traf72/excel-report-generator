namespace ExcelReporter.Interfaces.Providers
{
    public interface IDataItemValueProvider
    {
        object GetValue(string template, object dataItem);
    }
}