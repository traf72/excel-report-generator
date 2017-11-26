namespace ExcelReporter.Rendering.Providers.ParameterProviders
{
    public interface IParameterProvider
    {
        object GetParameterValue(string paramName);
    }
}