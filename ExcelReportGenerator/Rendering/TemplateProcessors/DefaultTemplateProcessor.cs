using ExcelReportGenerator.Exceptions;
using ExcelReportGenerator.Extensions;
using ExcelReportGenerator.Helpers;
using ExcelReportGenerator.Rendering.Providers;
using ExcelReportGenerator.Rendering.Providers.DataItemValueProviders;
using ExcelReportGenerator.Rendering.Providers.VariableProviders;
using System;

namespace ExcelReportGenerator.Rendering.TemplateProcessors
{
    /// <summary>
    /// Handles report templates
    /// </summary>
    public class DefaultTemplateProcessor : ITemplateProcessor
    {
        public DefaultTemplateProcessor(IPropertyValueProvider propertyValueProvider, IVariableValueProvider variableValueProvider,
            IMethodCallValueProvider methodCallValueProvider = null, IGenericDataItemValueProvider<HierarchicalDataItem> dataItemValueProvider = null)
        {
            PropertyValueProvider = propertyValueProvider ?? throw new ArgumentNullException(nameof(propertyValueProvider), ArgumentHelper.NullParamMessage);
            VariableValueProvider = variableValueProvider ?? throw new ArgumentNullException(nameof(variableValueProvider), ArgumentHelper.NullParamMessage);
            MethodCallValueProvider = methodCallValueProvider;
            DataItemValueProvider = dataItemValueProvider;
        }

        protected IPropertyValueProvider PropertyValueProvider { get; }

        protected IVariableValueProvider VariableValueProvider { get; }

        protected IGenericDataItemValueProvider<HierarchicalDataItem> DataItemValueProvider { get; }

        protected IMethodCallValueProvider MethodCallValueProvider { get; }

        public virtual string LeftTemplateBorder => "{";

        public virtual string RightTemplateBorder => "}";

        public virtual string MemberLabelSeparator => ":";

        public virtual string PropertyMemberLabel => "p";

        public virtual string MethodCallMemberLabel => "m";

        public virtual string DataItemMemberLabel => "di";

        public virtual string VariableMemberLabel => "v";

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

            string unwrappedTemplate = this.UnwrapTemplate(template.Trim()).Trim();
            int separatorIndex = unwrappedTemplate.IndexOf(MemberLabelSeparator, StringComparison.CurrentCultureIgnoreCase);
            if (separatorIndex == -1)
            {
                throw new InvalidTemplateException($"Invalid template \"{template}\". Cannot find separator \"{MemberLabelSeparator}\" between member label and member template");
            }

            string memberLabel = unwrappedTemplate.Substring(0, separatorIndex).ToLower().Trim();
            string memberTemplate = unwrappedTemplate.Substring(separatorIndex + MemberLabelSeparator.Length).Trim();
            if (memberLabel == PropertyMemberLabel)
            {
                return PropertyValueProvider.GetValue(memberTemplate);
            }
            if (memberLabel == DataItemMemberLabel)
            {
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
            if (memberLabel == MethodCallMemberLabel)
            {
                if (MethodCallValueProvider == null)
                {
                    throw new InvalidOperationException($"Template \"{template}\" contains method call but methodCallValueProvider is null");
                }
                return MethodCallValueProvider.CallMethod(memberTemplate, this, dataItem);
            }
            if (memberLabel == VariableMemberLabel)
            {
                return VariableValueProvider.GetVariable(memberTemplate);
            }

            throw new InvalidTemplateException($"Invalid template \"{template}\". Unknown member label \"{memberLabel}\"");
        }
    }
}