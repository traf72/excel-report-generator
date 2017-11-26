using System;
using System.Collections.Generic;
using ExcelReporter.Exceptions;
using ExcelReporter.Helpers;
using ExcelReporter.Rendering.Providers.DataItemValueProviders;

namespace ExcelReporter.Rendering.Providers.ParameterProviders
{
    /// <summary>
    /// Provides parameters values from dictionary
    /// </summary>
    public class DictionaryParameterProvider : IParameterProvider
    {
        private readonly IDataItemValueProvider _valueProvider = new DictionaryValueProvider<object>();
        private readonly IDictionary<string, object> _parameters;

        public DictionaryParameterProvider(IDictionary<string, object> parameters)
        {
            _parameters = parameters ?? throw new ArgumentNullException(nameof(parameters), ArgumentHelper.NullParamMessage);
        }

        public virtual object GetParameterValue(string paramName)
        {
            if (string.IsNullOrWhiteSpace(paramName))
            {
                throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(paramName));
            }

            try
            {
                return _valueProvider.GetValue(paramName, _parameters);
            }
            catch (KeyNotFoundException e)
            {
                throw new ParameterNotFoundException($"Cannot find paramater with name \"{paramName}\"", e);
            }
        }
    }
}