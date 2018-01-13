namespace ExcelReportGenerator.Rendering.Providers.VariableProviders
{
    public interface IVariableValueProvider
    {
        object GetVariable(string name);
    }
}