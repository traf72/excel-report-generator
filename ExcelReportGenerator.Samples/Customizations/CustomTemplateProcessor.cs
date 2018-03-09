using ExcelReportGenerator.Rendering;
using ExcelReportGenerator.Rendering.Providers;
using ExcelReportGenerator.Rendering.Providers.DataItemValueProviders;
using ExcelReportGenerator.Rendering.Providers.VariableProviders;
using ExcelReportGenerator.Rendering.TemplateProcessors;

namespace ExcelReportGenerator.Samples.Customizations
{
    public class CustomTemplateProcessor : DefaultTemplateProcessor
    {
        public CustomTemplateProcessor(IPropertyValueProvider propertyValueProvider, SystemVariableProvider systemVariableProvider, IMethodCallValueProvider methodCallValueProvider = null, IGenericDataItemValueProvider<HierarchicalDataItem> dataItemValueProvider = null) : base(propertyValueProvider, systemVariableProvider, methodCallValueProvider, dataItemValueProvider)
        {
        }

        public override string LeftTemplateBorder
        {
            get { return "<"; }
        }

        public override string RightTemplateBorder
        {
            get { return ">"; }
        }

        public override string MemberLabelSeparator
        {
            get { return "-"; }
        }

        public override string PropertyMemberLabel
        {
            get { return "prop"; }
        }

        public override string MethodCallMemberLabel
        {
            get { return "meth"; }
        }

        public override string DataItemMemberLabel
        {
            get { return "dataitem"; }
        }

        public override string SystemVariableMemberLabel
        {
            get { return "var"; }
        }

        public override string SystemFunctionMemberLabel
        {
            get { return "func"; }
        }

        public override string DataItemSelfTemplate
        {
            get { return "self"; }
        }

        public override string HorizontalPageBreakLabel
        {
            get { return "hbreak"; }
        }

        public override string VerticalPageBreakLabel
        {
            get { return "vbreak"; }
        }
    }
}