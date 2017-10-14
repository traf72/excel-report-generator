using JetBrains.Annotations;
using ReportEngine.Exceptions;
using ReportEngine.Interfaces.Providers;
using ReportEngine.Interfaces.TemplateProcessors;
using System;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace ReportEngine.Implementations.TemplateProcessors
{
    public class TemplateProcessor : ITemplateProcessor
    {
        private readonly IParameterProvider _parameterProvider;
        private readonly IDataItemValueProvider _dataItemValueProvider;
        private readonly IMethodContextProvider _methodContextProvider;
        private const string _typeValueSeparator = ":";
        private static readonly char[] _templateBorders = { '{', '}' };

        public TemplateProcessor([NotNull] IParameterProvider parameterProvider, IMethodContextProvider methodContextProvider = null,
            IDataItemValueProvider dataItemValueProvider = null)
        {
            if (parameterProvider == null)
            {
                throw new ArgumentNullException(nameof(parameterProvider), Constants.NullParamMessage);
            }

            _parameterProvider = parameterProvider;
            _methodContextProvider = methodContextProvider;
            _dataItemValueProvider = dataItemValueProvider;
        }

        public object GetValue(string template, object dataContext)
        {
            string templ = template.Trim(_templateBorders);
            int separatorIndex = templ.IndexOf(_typeValueSeparator, StringComparison.InvariantCulture);
            if (separatorIndex == -1)
            {
                throw new IncorrectTemplateException($"Incorrect template \"{template}\". Cannot find separator \"{_typeValueSeparator}\" between member type and member template");
            }

            string memberType = templ.Substring(0, separatorIndex).ToLower();
            string memberTemplate = templ.Substring(separatorIndex + 1);
            if (templ.StartsWith("p"))
            {
                return _parameterProvider.GetParameterValue(memberTemplate);
            }
            if (templ.StartsWith("di"))
            {
                // Значит это значение из элемента данных
                if (dataContext == null)
                {
                    throw new InvalidOperationException($"Template \"{template}\" contains data reference but dataContext is null");
                }
                if (_dataItemValueProvider == null)
                {
                    throw new InvalidOperationException($"Template \"{template}\" contains data reference but _dataItemValueProvider is null");
                }
                return _dataItemValueProvider.GetValue(memberTemplate, dataContext);
            }
            if (memberType == "fn")
            {
                // Значит это вызов метода
                if (_methodContextProvider == null)
                {
                    throw new InvalidOperationException($"Template \"{template}\" contains method call but _methodContextProvider is null");
                }

                return CallMethod(memberTemplate, dataContext);
            }

            throw new IncorrectTemplateException($"Incorrect template \"{template}\". Unknown member type \"{memberType}\"");
        }

        private object CallMethod(string methodTemplate, object dataContext)
        {
            //string[] parts = methodTemplate.Split(new[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
            Match match = Regex.Match(methodTemplate, @"(.*:)?(.+)\((.*)\)");
            string className = match.Groups[1].Value;
            string methodName = match.Groups[2].Value;
            string methodParams = match.Groups[3].Value;

            var context = _methodContextProvider.GetMethodContext(className);

            object[] callParams = methodParams
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(p => GetValue(p.Trim(), dataContext))
                .ToArray();

            MethodInfo method = context.GetType().GetMethod(methodName);
            return method.Invoke(context, callParams);
        }
    }
}