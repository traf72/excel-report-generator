using ExcelReporter.Exceptions;
using ExcelReporter.Interfaces.Providers;
using System;
using System.Collections.Generic;

namespace ExcelReporter.Implementations.Providers
{
    /// <summary>
    /// Provides parameters values from dictionary
    /// </summary>
    public class DictionaryParameterProvider : IParameterProvider
    {
        private readonly IDictionary<string, object> _parameters;

        public DictionaryParameterProvider(IDictionary<string, object> parameters)
        {
            _parameters = parameters ?? throw new ArgumentNullException(nameof(parameters), Constants.NullParamMessage);
        }

        public virtual object GetParameterValue(string paramName)
        {
            if (string.IsNullOrWhiteSpace(paramName))
            {
                throw new ArgumentException(Constants.EmptyStringParamMessage, nameof(paramName));
            }

            if (!_parameters.ContainsKey(paramName))
            {
                throw new ParameterNotFoundException($"Cannot find paramater with name \"{paramName}\"");
            }
            return _parameters[paramName];
        }
    }
}