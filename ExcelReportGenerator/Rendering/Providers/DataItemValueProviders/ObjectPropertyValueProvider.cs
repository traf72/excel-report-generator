using ExcelReportGenerator.Helpers;
using System;

namespace ExcelReportGenerator.Rendering.Providers.DataItemValueProviders
{
    // Provides properties values from object instance
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

        // Returns property value from data item object
        public virtual object GetValue(string propTemplate, object dataItem)
        {
            if (string.IsNullOrWhiteSpace(propTemplate))
            {
                throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(propTemplate));
            }

            return _reflectionHelper.GetValueOfPropertiesChain(propTemplate.Trim(), dataItem);
        }
    }
}