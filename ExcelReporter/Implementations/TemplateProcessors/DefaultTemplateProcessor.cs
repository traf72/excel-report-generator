using ExcelReporter.Exceptions;
using ExcelReporter.Interfaces.Providers;
using ExcelReporter.Interfaces.Providers.DataItemValueProviders;
using ExcelReporter.Interfaces.TemplateProcessors;
using System;

namespace ExcelReporter.Implementations.TemplateProcessors
{
    /// <summary>
    /// Handles report templates
    /// </summary>
    public class DefaultTemplateProcessor : IGenericTemplateProcessor<HierarchicalDataItem>
    {
        private const string TypeValueSeparator = ":";
        private const string LeftBorder = "{";
        private const string RightBorder = "}";

        private readonly IParameterProvider _parameterProvider;
        private readonly IGenericDataItemValueProvider<HierarchicalDataItem> _dataItemValueProvider;
        private readonly IMethodCallValueProvider _methodCallValueProvider;

        public DefaultTemplateProcessor(IParameterProvider parameterProvider, IMethodCallValueProvider methodCallValueProvider = null,
            IGenericDataItemValueProvider<HierarchicalDataItem> dataItemValueProvider = null)
        {
            if (parameterProvider == null)
            {
                throw new ArgumentNullException(nameof(parameterProvider), Constants.NullParamMessage);
            }

            _parameterProvider = parameterProvider;
            _methodCallValueProvider = methodCallValueProvider;
            _dataItemValueProvider = dataItemValueProvider;
        }

        public string LeftTemplateBorder => LeftBorder;

        public string RightTemplateBorder => RightBorder;

        public string TemplatePattern { get; } = $@"{LeftBorder}.+?{TypeValueSeparator}.+?{RightBorder}";

        /// <summary>
        /// Get value based on template
        /// </summary>
        /// <param name="dataItem">Data item that will be used if template is data item template</param>
        public virtual object GetValue(string template, HierarchicalDataItem dataItem = null)
        {
            if (string.IsNullOrWhiteSpace(template))
            {
                throw new ArgumentNullException(nameof(template), Constants.EmptyStringParamMessage);
            }

            string templ = template.Trim(LeftTemplateBorder[0], RightTemplateBorder[0], ' ');
            int separatorIndex = templ.IndexOf(TypeValueSeparator, StringComparison.InvariantCulture);
            if (separatorIndex == -1)
            {
                throw new IncorrectTemplateException($"Incorrect template \"{template}\". Cannot find separator \"{TypeValueSeparator}\" between member type and member template");
            }

            string memberType = templ.Substring(0, separatorIndex).ToLower().Trim();
            string memberTemplate = templ.Substring(separatorIndex + 1).Trim();
            if (templ.StartsWith("p"))
            {
                // Parameter value
                return _parameterProvider.GetParameterValue(memberTemplate);
            }
            if (templ.StartsWith("di"))
            {
                // Data item value
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
                // Method invocation
                if (_methodCallValueProvider == null)
                {
                    throw new InvalidOperationException($"Template \"{template}\" contains method call but methodCallValueProvider is null");
                }
                return _methodCallValueProvider.CallMethod(memberTemplate, this, dataItem, memberType == "ms");
            }

            throw new IncorrectTemplateException($"Incorrect template \"{template}\". Unknown member type \"{memberType}\"");
        }

        object ITemplateProcessor.GetValue(string template, object dataItem)
        {
            return GetValue(template, (HierarchicalDataItem)dataItem);
        }
    }
}