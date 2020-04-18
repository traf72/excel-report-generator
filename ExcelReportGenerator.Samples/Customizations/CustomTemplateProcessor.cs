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

        public override string LeftTemplateBorder => "<";

        public override string RightTemplateBorder => ">";

        public override string MemberLabelSeparator => "-";

        public override string PropertyMemberLabel => "prop";

        public override string MethodCallMemberLabel => "meth";

        public override string DataItemMemberLabel => "dataitem";

        public override string SystemVariableMemberLabel => "var";

        public override string SystemFunctionMemberLabel => "func";

        public override string DataItemSelfTemplate => "self";

        public override string HorizontalPageBreakLabel => "hbreak";

        public override string VerticalPageBreakLabel => "vbreak";
    }
}