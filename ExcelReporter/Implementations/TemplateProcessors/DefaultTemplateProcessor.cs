using ExcelReporter.Exceptions;
using ExcelReporter.Interfaces.Providers;
using ExcelReporter.Interfaces.TemplateProcessors;
using System;

namespace ExcelReporter.Implementations.TemplateProcessors
{
    public class DefaultTemplateProcessor : ITemplateProcessor
    {
        private readonly IParameterProvider _parameterProvider;
        private readonly IDataItemValueProvider _dataItemValueProvider;
        private readonly IMethodCallValueProvider _methodCallValueProvider;
        private const string _typeValueSeparator = ":";
        private static readonly char[] _templateBorders = { '{', '}' };

        public DefaultTemplateProcessor(IParameterProvider parameterProvider, IMethodCallValueProvider methodCallValueProvider = null,
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

        public string TemplatePattern { get; } = @"\{.+?:.+?\}";

        public virtual object GetValue(string template, HierarchicalDataItem dataItem)
        {
            if (string.IsNullOrWhiteSpace(template))
            {
                throw new ArgumentNullException(nameof(template), Constants.EmptyStringParamMessage);
            }

            string templ = template.Trim().Trim(_templateBorders).Trim();
            int separatorIndex = templ.IndexOf(_typeValueSeparator, StringComparison.InvariantCulture);
            if (separatorIndex == -1)
            {
                throw new IncorrectTemplateException($"Incorrect template \"{template}\". Cannot find separator \"{_typeValueSeparator}\" between member type and member template");
            }

            string memberType = templ.Substring(0, separatorIndex).ToLower().Trim();
            string memberTemplate = templ.Substring(separatorIndex + 1).Trim();
            if (templ.StartsWith("p"))
            {
                // Значит это значение параметра
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