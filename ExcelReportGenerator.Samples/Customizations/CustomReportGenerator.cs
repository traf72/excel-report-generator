using System;
using ExcelReportGenerator.Rendering;
using ExcelReportGenerator.Rendering.Providers;
using ExcelReportGenerator.Rendering.Providers.VariableProviders;
using ExcelReportGenerator.Rendering.TemplateProcessors;

namespace ExcelReportGenerator.Samples.Customizations
{
    public class CustomReportGenerator : DefaultReportGenerator
    {
        private IInstanceProvider _instanceProvider;
        private PanelParsingSettings _panelParsingSettings;
        private ITemplateProcessor _templateProcessor;

        public CustomReportGenerator(object report) : base(report)
        {
        }

        public override Type SystemFunctionsType => typeof(CustomSystemFunctions);

        public override SystemVariableProvider SystemVariableProvider => new CustomSystemVariableProvider();

        public override IInstanceProvider InstanceProvider => _instanceProvider ??= new CustomInstanceProvider(Report);

        public override ITemplateProcessor TemplateProcessor => _templateProcessor ??=
            new CustomTemplateProcessor(PropertyValueProvider, SystemVariableProvider, MethodCallValueProvider,
                DataItemValueProvider);

        public override PanelParsingSettings PanelParsingSettings
        {
            get
            {
                return _panelParsingSettings ??= new PanelParsingSettings
                {
                    PanelPrefixSeparator = "_",
                    SimplePanelPrefix = "simple",
                    DataSourcePanelPrefix = "data",
                    DynamicDataSourcePanelPrefix = "dynamic",
                    TotalsPanelPrefix = "total",
                    PanelPropertiesSeparators = new[] {","},
                    PanelPropertyNameValueSeparator = "="
                };
            }
        }
    }
}