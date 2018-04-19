using ClosedXML.Excel;
using ExcelReportGenerator.Attributes;
using ExcelReportGenerator.Converters;
using ExcelReportGenerator.Helpers;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelReportGenerator.Extensions;
using ExcelReportGenerator.License;

namespace ExcelReportGenerator.Rendering.Panels.ExcelPanels
{
    internal class ExcelPanelFactory
    {
        private readonly object _report;
        private readonly ITemplateProcessor _templateProcessor;
        private readonly PanelParsingSettings _panelParsingSettings;

        private IXLNamedRange _namedRange;
        private IDictionary<string, string> _properties;

        public ExcelPanelFactory(object report, ITemplateProcessor templateProcessor, PanelParsingSettings panelParsingSettings)
        {
            _report = report ?? throw new ArgumentNullException(nameof(report), ArgumentHelper.NullParamMessage);
            _templateProcessor = templateProcessor ?? throw new ArgumentNullException(nameof(templateProcessor), ArgumentHelper.NullParamMessage);
            _panelParsingSettings = panelParsingSettings ?? throw new ArgumentNullException(nameof(panelParsingSettings), ArgumentHelper.NullParamMessage);
            InitLicensing();
        }

        // This method is not refer to this class, it just hidden here
        public void InitLicensing()
        {
            new Licensing().LoadLicenseInfo();
        }

        public IExcelPanel Create(IXLNamedRange namedRange, IDictionary<string, string> properties)
        {
            _namedRange = namedRange ?? throw new ArgumentNullException(nameof(namedRange), ArgumentHelper.NullParamMessage);
            _properties = properties ?? new Dictionary<string, string>(0);

            IExcelPanel panel;

            int prefixIndex = namedRange.Name.IndexOf(_panelParsingSettings.PanelPrefixSeparator, StringComparison.CurrentCultureIgnoreCase);
            if (prefixIndex == -1)
            {
                throw new InvalidOperationException($"Panel name \"{namedRange.Name}\" does not contain prefix separator \"{_panelParsingSettings.PanelPrefixSeparator}\"");
            }

            string prefix = namedRange.Name.Substring(0, prefixIndex);
            if (prefix.Equals(_panelParsingSettings.SimplePanelPrefix, StringComparison.CurrentCultureIgnoreCase))
            {
                panel = CreateSimplePanel();
            }
            else if (prefix.Equals(_panelParsingSettings.DataSourcePanelPrefix, StringComparison.CurrentCultureIgnoreCase))
            {
                panel = CreateDataSourcePanel();
            }
            else if (prefix.Equals(_panelParsingSettings.DynamicDataSourcePanelPrefix, StringComparison.CurrentCultureIgnoreCase))
            {
                panel = CreateDynamicPanel();
            }
            else if (prefix.Equals(_panelParsingSettings.TotalsPanelPrefix, StringComparison.CurrentCultureIgnoreCase))
            {
                panel = CreateTotalsPanel();
            }
            else
            {
                throw new NotSupportedException($"Panel type with prefix \"{prefix}\" is not supported");
            }

            FillPanelProperties(panel);
            return panel;
        }

        private IExcelPanel CreateSimplePanel()
        {
            return new ExcelPanel(_namedRange.Ranges.First(), _report, _templateProcessor);
        }

        private IExcelPanel CreateDataSourcePanel()
        {
            return new ExcelDataSourcePanel(GetDataSourceProperty("Data source panel must have the property \"DataSource\""), _namedRange, _report, _templateProcessor);
        }

        private IExcelPanel CreateDynamicPanel()
        {
            return new ExcelDataSourceDynamicPanel(GetDataSourceProperty("Dynamic data source panel must have the property \"DataSource\""), _namedRange, _report, _templateProcessor);
        }

        private IExcelPanel CreateTotalsPanel()
        {
            return new ExcelTotalsPanel(GetDataSourceProperty("Totals panel must have the property \"DataSource\""), _namedRange, _report, _templateProcessor);
        }

        private string GetDataSourceProperty(string errorMessage)
        {
            if (_properties.TryGetValue("DataSource", out string dataSource))
            {
                return dataSource;
            }
            throw new InvalidOperationException(errorMessage);
        }

        private void FillPanelProperties(IExcelPanel panel)
        {
            if (!_properties.Any())
            {
                return;
            }

            PropertyInfo[] externalProperties = panel.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(p => p.IsDefined(typeof(ExternalPropertyAttribute), true)).ToArray();
            foreach (KeyValuePair<string, string> prop in _properties)
            {
                PropertyInfo externalProp = externalProperties.SingleOrDefault(p => p.Name == prop.Key);
                if (externalProp != null)
                {
                    var externalPropAttr = (ExternalPropertyAttribute)externalProp.GetCustomAttribute(typeof(ExternalPropertyAttribute));
                    externalProp.SetValue(panel, ConvertProperty(prop.Value, externalPropAttr.Converter), null);
                }
            }
        }

        private object ConvertProperty(string input, Type converter)
        {
            if (converter == null)
            {
                return input;
            }

            var converterInstance = (IConverter)Activator.CreateInstance(converter);
            return converterInstance.Convert(input);
        }
    }
}