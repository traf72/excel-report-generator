using JetBrains.Annotations;

namespace ReportEngine.Interfaces.Providers
{
    public interface IParameterProvider
    {
        object GetParameterValue([NotNull] string paramName);
    }
}