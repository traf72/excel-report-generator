using ExcelReporter.Exceptions;
using ExcelReporter.Interfaces.Providers;
using ExcelReporter.Interfaces.TemplateProcessors;
using System;

namespace ExcelReporter.Implementations.TemplateProcessors
{
    public class TemplateProcessor : ITemplateProcessor
    {
        private readonly IParameterProvider _parameterProvider;
        private readonly IDataItemValueProvider _dataItemValueProvider;
        private readonly IMethodCallValueProvider _methodCallValueProvider;
        private const string _typeValueSeparator = ":";
        private static readonly char[] _templateBorders = { '{', '}' };

        public TemplateProcessor(IParameterProvider parameterProvider, IMethodCallValueProvider methodCallValueProvider = null,
            IDataItemValueProvider dataItemValueProvider = null)
        {
            if (parameterProvider == null)
            {
                throw new ArgumentNullException(nameof(parameterProvider), Constants.NullParamMessage);
            }

            _parameterProvider = parameterProvider;
            _methodCallValueProvider = methodCallValueProvider;
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
            if (memberType == "m" || memberType == "ms")
            {
                // Значит это вызов метода
                if (_methodCallValueProvider == null)
                {
                    throw new InvalidOperationException($"Template \"{template}\" contains method call but methodCallValueProvider is null");
                }

                return _methodCallValueProvider.CallMethod(memberTemplate, this, dataItem, memberType == "ms");
            }

            throw new IncorrectTemplateException($"Incorrect template \"{template}\". Unknown member type \"{memberType}\"");
        }
    }
}