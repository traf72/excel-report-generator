using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Rendering
{
    [LicenceKeyPart(R = true)]
    public class PanelParsingSettings
    {
        public string PanelPrefixSeparator { get; set; }

        public string SimplePanelPrefix { get; set; }

        public string DataSourcePanelPrefix { get; set; }

        public string DynamicDataSourcePanelPrefix { get; set; }

        public string TotalsPanelPrefix { get; set; }

        public string[] PanelPropertiesSeparators { get; set; }

        public string PanelPropertyNameValueSeparator { get; set; }
    }
}