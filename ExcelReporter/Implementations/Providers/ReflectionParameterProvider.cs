using ExcelReporter.Exceptions;
using ExcelReporter.Interfaces.Providers;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelReporter.Implementations.Providers
{
    public class ReflectionParameterProvider : IParameterProvider
    {
        private readonly object _paramsContext;
        private readonly IDictionary<string, PropertyInfo> _propsCache = new Dictionary<string, PropertyInfo>();

        public ReflectionParameterProvider(object paramsContext)
        {
            if (paramsContext == null)
            {
                throw new ArgumentNullException(nameof(paramsContext), Constants.NullParamMessage);
            }
            _paramsContext = paramsContext;
        }

        public object GetParameterValue(string paramName)
        {
            if (string.IsNullOrWhiteSpace(paramName))
            {
                throw new ArgumentException(Constants.EmptyStringParamMessage, nameof(paramName));
            }

            if (_propsCache.ContainsKey(paramName))
            {
                return _propsCache[paramName].GetValue(_paramsContext);
            }

            // TODO Искать только свойства с атрибутом Parameter
            // TODO Также нужно искать и в родительских классах
            PropertyInfo prop = _paramsContext.GetType().GetProperty(paramName);
            if (prop == null)
            {
                throw new ParameterNotFoundException($"Cannot find paramater with name \"{paramName}\" in class \"{_paramsContext.GetType().Name}\" and its parents");
            }
            _propsCache[paramName] = prop;
            return _propsCache[paramName].GetValue(_paramsContext);
        }
    }
}