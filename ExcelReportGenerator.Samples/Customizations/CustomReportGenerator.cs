using ExcelReportGenerator.Rendering;
using ExcelReportGenerator.Rendering.Providers;
using ExcelReportGenerator.Rendering.Providers.VariableProviders;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using System;

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

        public override Type SystemFunctionsType
        {
            get { return typeof(CustomSystemFunctions); }
        }

        public override SystemVariableProvider SystemVariableProvider
        {
            get { return new CustomSystemVariableProvider(); }
        }

        public override IInstanceProvider InstanceProvider
        {
            get { return _instanceProvider ?? (_instanceProvider = new CustomInstanceProvider(_report)); }
        }

        public override ITemplateProcessor TemplateProcessor
        {
            get
            {
                return _templateProcessor ?? (_templateProcessor = new CustomTemplateProcessor(PropertyValueProvider, SystemVariableProvider, MethodCallValueProvider, DataItemValueProvider));
            }
        }

        public override PanelParsingSettings PanelParsingSettings
        {
            get
            {
                return _panelParsingSettings ?? (_panelParsingSettings = new PanelParsingSettings
                {
                    PanelPrefixSeparator = "_",
                    SimplePanelPrefix = "simple",
                    DataSourcePanelPrefix = "data",
                    DynamicDataSourcePanelPrefix = "dynamic",
                    TotalsPanelPrefix = "total",
                    PanelPropertiesSeparators = new[] { "," },
                    PanelPropertyNameValueSeparator = "=",
                });
            }
        }
    }
}