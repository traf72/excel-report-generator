using ExcelReportGenerator.Exceptions;
using ExcelReportGenerator.Extensions;
using ExcelReportGenerator.Helpers;
using ExcelReportGenerator.Rendering.Providers;
using ExcelReportGenerator.Rendering.Providers.DataItemValueProviders;
using ExcelReportGenerator.Rendering.Providers.VariableProviders;
using System;
using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Rendering.TemplateProcessors
{
    /// <summary>
    /// Default implementation of <see cref="ITemplateProcessor" /> 
    /// </summary>
    [LicenceKeyPart(U = true)]
    public class DefaultTemplateProcessor : ITemplateProcessor
    {
        /// <param name="propertyValueProvider">See <see cref="IPropertyValueProvider" /></param>
        /// <param name="systemVariableProvider">See <see cref="Providers.VariableProviders.SystemVariableProvider" /></param>
        /// <param name="methodCallValueProvider">See <see cref="IMethodCallValueProvider" /></param>
        /// <param name="dataItemValueProvider">See <see cref="IGenericDataItemValueProvider{T}" /></param>
        public DefaultTemplateProcessor(IPropertyValueProvider propertyValueProvider, SystemVariableProvider systemVariableProvider,
            IMethodCallValueProvider methodCallValueProvider = null, IGenericDataItemValueProvider<HierarchicalDataItem> dataItemValueProvider = null)
        {
            PropertyValueProvider = propertyValueProvider ?? throw new ArgumentNullException(nameof(propertyValueProvider), ArgumentHelper.NullParamMessage);
            SystemVariableProvider = systemVariableProvider ?? throw new ArgumentNullException(nameof(systemVariableProvider), ArgumentHelper.NullParamMessage);
            MethodCallValueProvider = methodCallValueProvider;
            DataItemValueProvider = dataItemValueProvider;
        }

        /// <summary>
        /// See <see cref="IPropertyValueProvider" /> 
        /// </summary>
        protected IPropertyValueProvider PropertyValueProvider { get; }

        /// <summary>
        /// See <see cref="Providers.VariableProviders.SystemVariableProvider" /> 
        /// </summary>
        protected SystemVariableProvider SystemVariableProvider { get; }

        /// <summary>
        /// See <see cref="IGenericDataItemValueProvider{T}" /> 
        /// </summary>
        protected IGenericDataItemValueProvider<HierarchicalDataItem> DataItemValueProvider { get; }

        /// <summary>
        /// See <see cref="IMethodCallValueProvider" /> 
        /// </summary>
        protected IMethodCallValueProvider MethodCallValueProvider { get; }

        internal Type SystemFunctionsType { get; set; } = typeof(SystemFunctions);

        /// <inheritdoc />
        public virtual string LeftTemplateBorder => "{";

        /// <inheritdoc />
        public virtual string RightTemplateBorder => "}";

        /// <inheritdoc />
        public virtual string MemberLabelSeparator => ":";

        /// <inheritdoc />
        public virtual string PropertyMemberLabel => "p";

        /// <inheritdoc />
        public virtual string MethodCallMemberLabel => "m";

        /// <inheritdoc />
        public virtual string DataItemMemberLabel => "di";

        /// <inheritdoc />
        public virtual string SystemVariableMemberLabel => "sv";

        /// <inheritdoc />
        public virtual string SystemFunctionMemberLabel => "sf";

        /// <inheritdoc />
        public virtual string HorizontalPageBreakLabel => "HorizPageBreak";

        /// <inheritdoc />
        public virtual string VerticalPageBreakLabel => "VertPageBreak";

        /// <inheritdoc />
        public virtual string DataItemSelfTemplate => DataItemMemberLabel;

        /// <inheritdoc />
        /// <exception cref="ArgumentException">Thrown when <paramref name="template" /> is null, empty string or whitespace</exception>
        /// <exception cref="InvalidTemplateException"></exception>
        /// <exception cref="InvalidOperationException"></exception>
        public virtual object GetValue(string template, HierarchicalDataItem dataItem = null)
        {
            if (string.IsNullOrWhiteSpace(template))
            {
                throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(template));
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
            if (memberLabel == SystemVariableMemberLabel)
            {
                return SystemVariableProvider.GetVariable(memberTemplate);
            }
            if (memberLabel == SystemFunctionMemberLabel)
            {
                if (SystemFunctionsType == null)
                {
                    throw new InvalidOperationException($"Template \"{template}\" contains system function call but property SystemFunctionsType is null");
                }
                if (MethodCallValueProvider == null)
                {
                    throw new InvalidOperationException($"Template \"{template}\" contains system function call but methodCallValueProvider is null");
                }
                return MethodCallValueProvider.CallMethod(memberTemplate, SystemFunctionsType, this, dataItem);
            }

            throw new InvalidTemplateException($"Invalid template \"{template}\". Unknown member label \"{memberLabel}\"");
        }
    }
}