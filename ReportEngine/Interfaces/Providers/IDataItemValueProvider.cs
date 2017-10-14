using JetBrains.Annotations;

namespace ReportEngine.Interfaces.Providers
{
    public interface IDataItemValueProvider
    {
        object GetValue([NotNull] string template, [NotNull] object dataItem);
    }
}