namespace ExcelReporter.Rendering.Providers
{
    public interface IParameterProvider
    {
        object GetParameterValue(string paramName);
    }
}