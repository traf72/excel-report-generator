using ExcelReporter.Exceptions;
using ExcelReporter.Extensions;
using ExcelReporter.Helpers;
using ExcelReporter.Rendering.Providers;
using ExcelReporter.Rendering.Providers.DataItemValueProviders;
using System;

namespace ExcelReporter.Rendering.TemplateProcessors
{
    /// <summary>
    /// Handles report templates
    /// </summary>
    public class DefaultTemplateProcessor : IGenericTemplateProcessor<HierarchicalDataItem>
    {
        private const string TypeValueSeparator = ":";

        public DefaultTemplateProcessor(IPropertyValueProvider propertyValueProvider, IMethodCallValueProvider methodCallValueProvider = null,
            IGenericDataItemValueProvider<HierarchicalDataItem> dataItemValueProvider = null)
        {
            PropertyValueProvider = propertyValueProvider ?? throw new ArgumentNullException(nameof(propertyValueProvider), ArgumentHelper.NullParamMessage);
            MethodCallValueProvider = methodCallValueProvider;
            DataItemValueProvider = dataItemValueProvider;
        }

        protected IPropertyValueProvider PropertyValueProvider { get; }

        protected IGenericDataItemValueProvider<HierarchicalDataItem> DataItemValueProvider { get; }

        protected IMethodCallValueProvider MethodCallValueProvider { get; }

        // TODO Обязательно протестировать переопределение границ (в том числе на границы с более, чем одним символом)
        public virtual string LeftTemplateBorder => "{";

        // TODO Обязательно протестировать переопределение границ (в том числе на границы с более, чем одним символом)
        public virtual string RightTemplateBorder => "}";

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
                // Property of field value
                return PropertyValueProvider.GetValue(memberTemplate);
            }
            if (unwrappedTemplate.StartsWith("di"))
            {
                // Data item value
                if (dataItem == null)
                {
                    throw new InvalidOperationException($"Template \"{template}\" contains data reference but dataItem is null");
                }
                if (DataItemValueProvider == null)
                {
                    throw new InvalidOperationException($"Template \"{template}\" contains data reference but dataItemValueProvider is null");
                }
                return DataItemValueProvider.GetValue(memberTemplate, dataItem);
            }
            if (memberType == "m")
            {
                // Method invocation
                if (MethodCallValueProvider == null)
                {
                    throw new InvalidOperationException($"Template \"{template}\" contains method call but methodCallValueProvider is null");
                }
                return MethodCallValueProvider.CallMethod(memberTemplate, this, dataItem);
            }

            throw new IncorrectTemplateException($"Incorrect template \"{template}\". Unknown member type \"{memberType}\"");
        }

        object ITemplateProcessor.GetValue(string template, object dataItem)
        {
            return GetValue(template, (HierarchicalDataItem)dataItem);
        }
    }
}