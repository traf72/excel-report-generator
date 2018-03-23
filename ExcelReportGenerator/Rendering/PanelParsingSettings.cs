using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Rendering
{
    /// <summary>
    /// Allow to configure panel parsing settings
    /// </summary>
    [LicenceKeyPart(R = true)]
    public class PanelParsingSettings
    {
        /// <summary>
        /// Separator between panel type prefix and panel name
        /// Cannot be null, empty string or whitespace
        /// </summary>
        /// <example>For "s_panel" panel name the separator is "_"</example>
        public string PanelPrefixSeparator { get; set; }

        /// <summary>
        /// Prefix for simple panel
        /// Cannot be null, empty string or whitespace
        /// </summary>
        /// <example>For "s_header" panel name the prefix is "s"</example>
        public string SimplePanelPrefix { get; set; }

        /// <summary>
        /// Prefix for data source panel
        /// Cannot be null, empty string or whitespace
        /// </summary>
        /// <example>For "d_data" panel name the prefix is "d"</example>
        public string DataSourcePanelPrefix { get; set; }

        /// <summary>
        /// Prefix for dynamic panel
        /// Cannot be null, empty string or whitespace
        /// </summary>
        /// <example>For "dyn_data" panel name the prefix is "dyn"</example>
        public string DynamicDataSourcePanelPrefix { get; set; }

        /// <summary>
        /// Prefix for totals panel
        /// Cannot be null, empty string or whitespace
        /// </summary>
        /// <example>For "t_data" panel name the prefix is "t"</example>
        public string TotalsPanelPrefix { get; set; }

        /// <summary>
        /// Separator between different panel properties. Allow to specify multiple separators
        /// Cannot be null or empty array
        /// </summary>
        public string[] PanelPropertiesSeparators { get; set; }

        /// <summary>
        /// Separator between panel property name and value
        /// Cannot be null, empty string or whitespace
        /// </summary>
        public string PanelPropertyNameValueSeparator { get; set; }
    }
}