using ExcelReporter.Exceptions;
using ExcelReporter.Interfaces.Providers;
using ExcelReporter.Interfaces.TemplateProcessors;
using System;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace ExcelReporter.Implementations.TemplateProcessors
{
    public class TemplateProcessor : ITemplateProcessor
    {
        private readonly IParameterProvider _parameterProvider;
        private readonly IDataItemValueProvider _dataItemValueProvider;
        private readonly IMethodContextProvider _methodContextProvider;
        private const string _typeValueSeparator = ":";
        private static readonly char[] _templateBorders = { '{', '}' };

        public TemplateProcessor(IParameterProvider parameterProvider, IMethodContextProvider methodContextProvider = null,
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

        public string Pattern { get; } = @"\{.+?:.+?\}";

        public object GetValue(string template, HierarchicalDataItem dataItem)
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
                if (dataItem == null)
                {
                    throw new InvalidOperationException($"Template \"{template}\" contains data reference but dataItem is null");
                }
                if (_dataItemValueProvider == null)
                {
                    throw new InvalidOperationException($"Template \"{template}\" contains data reference but dataItemValueProvider is null");
                }
                return _dataItemValueProvider.GetValue(memberTemplate, dataItem);
            }
            if (memberType == "fn")
            {
                // Значит это вызов метода
                if (_methodContextProvider == null)
                {
                    throw new InvalidOperationException($"Template \"{template}\" contains method call but methodContextProvider is null");
                }

                return CallMethod(memberTemplate, dataItem);
            }

            throw new IncorrectTemplateException($"Incorrect template \"{template}\". Unknown member type \"{memberType}\"");
        }

        private object CallMethod(string methodTemplate, HierarchicalDataItem dataItem)
        {
            //string[] parts = methodTemplate.Split(new[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
            Match match = Regex.Match(methodTemplate, @"(.*:)?(.+)\((.*)\)");
            string className = match.Groups[1].Value;
            string methodName = match.Groups[2].Value;
            string methodParams = match.Groups[3].Value;

            object context = _methodContextProvider.GetMethodContext(className.Trim(':'));

            object[] callParams = methodParams
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(p => GetValue(p.Trim(), dataItem))
                .ToArray();

            MethodInfo method = context.GetType().GetMethod(methodName);
            return method.Invoke(context, callParams);
        }
    }
}