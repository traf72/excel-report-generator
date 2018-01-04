using ClosedXML.Excel;
using ExcelReporter.Helpers;
using ExcelReporter.Rendering.Panels.ExcelPanels;
using ExcelReporter.Rendering.Parsers;
using ExcelReporter.Rendering.Providers;
using ExcelReporter.Rendering.Providers.DataItemValueProviders;
using ExcelReporter.Rendering.TemplateProcessors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelReporter.Rendering
{
    public class DefaultReportGenerator
    {
        protected readonly object _report;

        private ITypeProvider _typeProvider;
        private IInstanceProvider _instanceProvider;
        private IPropertyValueProvider _propertyValueProvider;
        private IMethodCallValueProvider _methodCallValueProvider;
        private IGenericDataItemValueProvider<HierarchicalDataItem> _dataItemValueProvider;
        private ITemplateProcessor _templateProcessor;
        private IPanelPropertiesParser _panelPropertiesParser;
        private PanelParsingSettings _panelParsingSettings;
        private string _panelsRegexPattern;

        public DefaultReportGenerator(object report)
        {
            _report = report ?? throw new ArgumentNullException(nameof(report), ArgumentHelper.NullParamMessage);
        }

        public virtual ITypeProvider TypeProvider => _typeProvider ?? (_typeProvider = new DefaultTypeProvider(defaultType: _report.GetType()));

        public virtual IInstanceProvider InstanceProvider => _instanceProvider ?? (_instanceProvider = new DefaultInstanceProvider(_report));

        public virtual IPropertyValueProvider PropertyValueProvider => _propertyValueProvider ?? (_propertyValueProvider = new DefaultPropertyValueProvider(TypeProvider, InstanceProvider));

        public virtual IMethodCallValueProvider MethodCallValueProvider => _methodCallValueProvider ?? (_methodCallValueProvider = new DefaultMethodCallValueProvider(TypeProvider, InstanceProvider));

        public virtual IGenericDataItemValueProvider<HierarchicalDataItem> DataItemValueProvider => _dataItemValueProvider ?? (_dataItemValueProvider = new HierarchicalDataItemValueProvider());

        public virtual ITemplateProcessor TemplateProcessor => _templateProcessor ?? (_templateProcessor = new DefaultTemplateProcessor(PropertyValueProvider, MethodCallValueProvider, DataItemValueProvider));

        public virtual IPanelPropertiesParser PanelPropertiesParser => _panelPropertiesParser ?? (_panelPropertiesParser = new DefaultPanelPropertiesParser(PanelParsingSettings));

        public virtual PanelParsingSettings PanelParsingSettings => _panelParsingSettings ?? (_panelParsingSettings = new PanelParsingSettings
        {
            PanelPrefixSeparator = "_",
            SimplePanelPrefix = "s",
            DataSourcePanelPrefix = "d",
            DynamicDataSourcePanelPrefix = "dyn",
            TotalsPanelPrefix = "t",
            PanelPropertiesSeparators = new[] { Environment.NewLine, "\t", ";" },
            PanelPropertyNameValueSeparator = "=",
        });

        private string PanelsRegexPattern
        {
            get
            {
                if (_panelsRegexPattern == null)
                {
                    string patternPanelsPrefixes = $"{Regex.Escape(PanelParsingSettings.SimplePanelPrefix)}" +
                                                   $"|{Regex.Escape(PanelParsingSettings.DataSourcePanelPrefix)}" +
                                                   $"|{Regex.Escape(PanelParsingSettings.DynamicDataSourcePanelPrefix)}" +
                                                   $"|{Regex.Escape(PanelParsingSettings.TotalsPanelPrefix)}";

                    _panelsRegexPattern = $"^({patternPanelsPrefixes}){Regex.Escape(PanelParsingSettings.PanelPrefixSeparator)}.+$";
                }

                return _panelsRegexPattern;
            }
        }

        public XLWorkbook Render(XLWorkbook reportTemplate, IXLWorksheet[] worksheets = null)
        {
            if (reportTemplate == null)
            {
                throw new ArgumentNullException(nameof(reportTemplate), ArgumentHelper.NullParamMessage);
            }

            if (worksheets == null || !worksheets.Any())
            {
                worksheets = reportTemplate.Worksheets.ToArray();
            }

            if (!worksheets.Any())
            {
                return reportTemplate;
            }

            IList<IXLNamedRange> workbookPanels = GetPanelsNamedRanges(reportTemplate.NamedRanges);
            foreach (IXLWorksheet ws in worksheets)
            {
                IList<IXLNamedRange> worksheetPanels = GetPanelsNamedRanges(ws.NamedRanges);
                foreach (IXLNamedRange workbookPanel in workbookPanels)
                {
                    if (workbookPanel.Ranges.First().Worksheet == ws
                        && !worksheetPanels.Any(p => p.Name.Trim().Equals(workbookPanel.Name.Trim(), StringComparison.CurrentCultureIgnoreCase)))
                    {
                        worksheetPanels.Add(workbookPanel);
                    }
                }

                IDictionary<string, (IExcelPanel, string)> panelsFlatView = GetPanelsFlatView(worksheetPanels);
                IExcelPanel rootPanel = new ExcelPanel(ws.Range(ws.FirstCellUsed(), ws.LastCellUsed()), _report, _templateProcessor);
                MakePanelsHierarchy(panelsFlatView, rootPanel);
                rootPanel.Render();
            }

            return reportTemplate;
        }

        private IList<IXLNamedRange> GetPanelsNamedRanges(IXLNamedRanges namedRanges)
        {
            return namedRanges.Where(r => Regex.IsMatch(r.Name, PanelsRegexPattern, RegexOptions.IgnoreCase)).ToList();
        }

        private IDictionary<string, (IExcelPanel, string)> GetPanelsFlatView(IEnumerable<IXLNamedRange> panelsNamedRanges)
        {
            var panelFactory = new ExcelPanelFactory(_report, TemplateProcessor, PanelParsingSettings);
            IDictionary<string, (IExcelPanel, string)> panels = new Dictionary<string, (IExcelPanel, string)>();
            foreach (IXLNamedRange namedRange in panelsNamedRanges)
            {
                IDictionary<string, string> panelProperties = PanelPropertiesParser.Parse(namedRange.Comment);
                IExcelPanel panel = panelFactory.Create(namedRange, panelProperties);
                panelProperties.TryGetValue("ParentPanel", out string parentPanelName);
                panels[namedRange.Name] = (panel, parentPanelName);
            }
            return panels;
        }

        private void MakePanelsHierarchy(IDictionary<string, (IExcelPanel, string)> panelsFlatView, IExcelPanel rootPanel)
        {
            foreach (KeyValuePair<string, (IExcelPanel, string)> panelFlat in panelsFlatView)
            {
                (IExcelPanel panel, string parentPanelName) = panelFlat.Value;
                if (string.IsNullOrWhiteSpace(parentPanelName))
                {
                    rootPanel.Children.Add(panel);
                    panel.Parent = rootPanel;
                    continue;
                }

                if (panelsFlatView.ContainsKey(parentPanelName))
                {
                    (IExcelPanel parentPanel, _) = panelsFlatView[parentPanelName];
                    parentPanel.Children.Add(panel);
                    panel.Parent = parentPanel;
                }
                else
                {
                    throw new InvalidOperationException($"Cannot find parent panel with name \"{parentPanelName}\" for panel \"{panelFlat.Key}\"");
                }
            }
        }
    }
}