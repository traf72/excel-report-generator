namespace ExcelReporter.Interfaces.Providers
{
    public interface IParameterProvider
    {
        object GetParameterValue(string paramName);
    }
}