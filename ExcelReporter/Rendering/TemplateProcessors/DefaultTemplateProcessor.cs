using ExcelReporter.Exceptions;
using ExcelReporter.Extensions;
using ExcelReporter.Helpers;
using ExcelReporter.Rendering.Providers;
using ExcelReporter.Rendering.Providers.DataItemValueProviders;
using System;
using ExcelReporter.Rendering.Providers.ParameterProviders;

namespace ExcelReporter.Rendering.TemplateProcessors
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
            _parameterProvider = parameterProvider ?? throw new ArgumentNullException(nameof(parameterProvider), ArgumentHelper.NullParamMessage);
            _methodCallValueProvider = methodCallValueProvider;
            _dataItemValueProvider = dataItemValueProvider;
        }

        // TODO Обязательно протестировать переопределение границ (в том числе на границы с более, чем одним символом)
        public virtual string LeftTemplateBorder => LeftBorder;

        // TODO Обязательно протестировать переопределение границ (в том числе на границы с более, чем одним символом)
        public virtual string RightTemplateBorder => RightBorder;

        public string TemplatePattern => $@"{LeftTemplateBorder}.+?{TypeValueSeparator}.+?{RightTemplateBorder}";

        /// <summary>
        /// Get value based on template
        /// </summary>
        /// <param name="dataItem">Data item that will be used if template is data item template</param>
        public virtual object GetValue(string template, HierarchicalDataItem dataItem = null)
        {
            if (string.IsNullOrWhiteSpace(template))
            {
                throw new ArgumentNullException(nameof(template), ArgumentHelper.EmptyStringParamMessage);
            }

            string unwrappedTemplate = this.UnwrapTemplate(template).Trim();
            int separatorIndex = unwrappedTemplate.IndexOf(TypeValueSeparator, StringComparison.InvariantCulture);
            if (separatorIndex == -1)
            {
                throw new IncorrectTemplateException($"Incorrect template \"{template}\". Cannot find separator \"{TypeValueSeparator}\" between member type and member template");
            }

            string memberType = unwrappedTemplate.Substring(0, separatorIndex).ToLower().Trim();
            string memberTemplate = unwrappedTemplate.Substring(separatorIndex + 1).Trim();
            if (unwrappedTemplate.StartsWith("p"))
            {
                // Parameter value
                return _parameterProvider.GetParameterValue(memberTemplate);
            }
            if (unwrappedTemplate.StartsWith("di"))
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
            if (memberType == "m")
            {
                // Method invocation
                if (_methodCallValueProvider == null)
                {
                    throw new InvalidOperationException($"Template \"{template}\" contains method call but methodCallValueProvider is null");
                }
                return _methodCallValueProvider.CallMethod(memberTemplate, this, dataItem);
            }

            throw new IncorrectTemplateException($"Incorrect template \"{template}\". Unknown member type \"{memberType}\"");
        }

        object ITemplateProcessor.GetValue(string template, object dataItem)
        {
            return GetValue(template, (HierarchicalDataItem)dataItem);
        }
    }
}