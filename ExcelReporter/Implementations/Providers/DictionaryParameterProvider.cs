using System;
using ExcelReporter.Exceptions;
using ExcelReporter.Interfaces.Providers;
using System.Collections.Generic;

namespace ExcelReporter.Implementations.Providers
{
    public struct DictionaryParameterProvider : IParameterProvider
    {
        private readonly IDictionary<string, object> _parameters;

        public DictionaryParameterProvider(IDictionary<string, object> parameters)
        {
            _parameters = parameters;
        }

        public object GetParameterValue(string paramName)
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