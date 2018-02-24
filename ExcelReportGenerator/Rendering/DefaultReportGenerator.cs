using ClosedXML.Excel;
using ExcelReportGenerator.Excel;
using ExcelReportGenerator.Helpers;
using ExcelReportGenerator.Rendering.EventArgs;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Rendering.Parsers;
using ExcelReportGenerator.Rendering.Providers;
using ExcelReportGenerator.Rendering.Providers.DataItemValueProviders;
using ExcelReportGenerator.Rendering.Providers.VariableProviders;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelReportGenerator.Rendering
{
    public class DefaultReportGenerator
    {
        public event EventHandler<ReportRenderEventArgs> BeforeReportRender;

        public event EventHandler<WorksheetRenderEventArgs> BeforeWorksheetRender;

        public event EventHandler<WorksheetRenderEventArgs> AfterWorksheetRender;

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

        public virtual Type SystemFunctionsType { get; set; } = typeof(SystemFunctions);

        public virtual SystemVariableProvider SystemVariableProvider { get; set; } = new SystemVariableProvider();

        public virtual ITypeProvider TypeProvider => _typeProvider ?? (_typeProvider = new DefaultTypeProvider(defaultType: _report.GetType()));

        public virtual IInstanceProvider InstanceProvider => _instanceProvider ?? (_instanceProvider = new DefaultInstanceProvider(_report));

        public virtual IPropertyValueProvider PropertyValueProvider => _propertyValueProvider ?? (_propertyValueProvider = new DefaultPropertyValueProvider(TypeProvider, InstanceProvider));

        public virtual IMethodCallValueProvider MethodCallValueProvider => _methodCallValueProvider ?? (_methodCallValueProvider = new DefaultMethodCallValueProvider(TypeProvider, InstanceProvider));

        public virtual IGenericDataItemValueProvider<HierarchicalDataItem> DataItemValueProvider => _dataItemValueProvider ?? (_dataItemValueProvider = new HierarchicalDataItemValueProvider());

        public virtual ITemplateProcessor TemplateProcessor
        {
            get
            {
                if (_templateProcessor == null)
                {
                    _templateProcessor = new DefaultTemplateProcessor(PropertyValueProvider, SystemVariableProvider, MethodCallValueProvider, DataItemValueProvider)
                    {
                        SystemFunctionsType = SystemFunctionsType
                    };
                    if (DataItemValueProvider is HierarchicalDataItemValueProvider p && string.IsNullOrWhiteSpace(p.DataItemSelfTemplate))
                    {
                        p.DataItemSelfTemplate = _templateProcessor.DataItemMemberLabel;
                    }
                }

                return _templateProcessor;
            }
        }

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

        private IPanelPropertiesParser PanelPropertiesParser => _panelPropertiesParser ?? (_panelPropertiesParser = new DefaultPanelPropertiesParser(PanelParsingSettings));

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
            if (SystemVariableProvider == null)
            {
                throw new Exception($"Property {nameof(SystemVariableProvider)} cannot be null");
            }

            SystemVariableProvider.RenderDate = DateTime.Now;

            OnBeforeReportRender(new ReportRenderEventArgs { Workbook = reportTemplate });

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
                SystemVariableProvider.SheetName = ws.Name;
                SystemVariableProvider.SheetNumber = ws.Position;

                OnBeforeWorksheetRender(new WorksheetRenderEventArgs { Worksheet = ws });

                if (!ws.CellsUsed().Any())
                {
                    continue;
                }

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
                IExcelPanel rootPanel = new ExcelPanel(GetRootRange(ws, worksheetPanels), _report, TemplateProcessor);
                MakePanelsHierarchy(panelsFlatView, rootPanel);
                rootPanel.Render();

                OnAfterWorksheetRender(new WorksheetRenderEventArgs { Worksheet = ws });
            }

            return reportTemplate;
        }

        private IXLRange GetRootRange(IXLWorksheet ws, IList<IXLNamedRange> worksheetPanels)
        {
            return ws.Range(ws.FirstCell(), GetRootRangeLastCell(ws.LastCellUsed(), worksheetPanels ?? new List<IXLNamedRange>()));
        }

        private IXLCell GetRootRangeLastCell(IXLCell lastWorksheetCellUsed, IList<IXLNamedRange> worksheetPanels)
        {
            return ExcelHelper.GetMaxCell(worksheetPanels.SelectMany(p => p.Ranges.First().Cells()).Concat(new[] { lastWorksheetCellUsed }).ToArray());
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

        protected virtual void OnBeforeReportRender(ReportRenderEventArgs e)
        {
            BeforeReportRender?.Invoke(this, e);
        }

        protected virtual void OnBeforeWorksheetRender(WorksheetRenderEventArgs e)
        {
            BeforeWorksheetRender?.Invoke(this, e);
        }

        protected virtual void OnAfterWorksheetRender(WorksheetRenderEventArgs e)
        {
            AfterWorksheetRender?.Invoke(this, e);
        }
    }
}