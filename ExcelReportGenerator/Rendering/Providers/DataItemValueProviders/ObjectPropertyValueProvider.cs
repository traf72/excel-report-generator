using ExcelReportGenerator.Helpers;

namespace ExcelReportGenerator.Rendering.Providers.DataItemValueProviders;

/// <summary>
/// Provides properties values from object instance
/// </summary>
internal class ObjectPropertyValueProvider : IDataItemValueProvider
{
    private readonly IReflectionHelper _reflectionHelper;

    public ObjectPropertyValueProvider() : this(new ReflectionHelper())
    {
    }

    internal ObjectPropertyValueProvider(IReflectionHelper reflectionHelper)
    {
        _reflectionHelper = reflectionHelper;
    }

    /// <summary>
    /// Returns property value from data item object
    /// </summary>
    public virtual object GetValue(string propTemplate, object dataItem)
    {
        if (string.IsNullOrWhiteSpace(propTemplate))
        {
            throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(propTemplate));
        }

        return _reflectionHelper.GetValueOfPropertiesChain(propTemplate.Trim(), dataItem);
    }
}