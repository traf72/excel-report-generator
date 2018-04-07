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
    /// <summary>
    /// Report generator with default configurations
    /// </summary>
    public class DefaultReportGenerator
    {
        /// <summary>
        /// Event raised before render report 
        /// </summary>
        public event EventHandler<ReportRenderEventArgs> BeforeReportRender;

        /// <summary>
        /// Event raised before render each worksheet
        /// </summary>
        public event EventHandler<WorksheetRenderEventArgs> BeforeWorksheetRender;

        /// <summary>
        /// Event raised after render each worksheet
        /// </summary>
        public event EventHandler<WorksheetRenderEventArgs> AfterWorksheetRender;

        /// <summary>
        /// Report object
        /// </summary>
        protected readonly object Report;

        private ITypeProvider _typeProvider;
        private IInstanceProvider _instanceProvider;
        private IPropertyValueProvider _propertyValueProvider;
        private IMethodCallValueProvider _methodCallValueProvider;
        private IGenericDataItemValueProvider<HierarchicalDataItem> _dataItemValueProvider;
        private ITemplateProcessor _templateProcessor;
        private IPanelPropertiesParser _panelPropertiesParser;
        private PanelParsingSettings _panelParsingSettings;
        private string _panelsRegexPattern;

        /// <param name="report">Report object</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="report"/> is null</exception>
        public DefaultReportGenerator(object report)
        {
            Report = report ?? throw new ArgumentNullException(nameof(report), ArgumentHelper.NullParamMessage);
        }

        /// <summary>
        /// Type of system functions. Default value is <see cref="SystemFunctions"/>
        /// </summary>
        public virtual Type SystemFunctionsType { get; set; } = typeof(SystemFunctions);

        /// <summary>
        /// System variable provider. Default value is instance of <see cref="Providers.VariableProviders.SystemVariableProvider"/>
        /// </summary>
        public virtual SystemVariableProvider SystemVariableProvider { get; set; } = new SystemVariableProvider();

        /// <summary>
        /// Type provider. Default value is instance of <see cref="DefaultTypeProvider"/>
        /// </summary>
        public virtual ITypeProvider TypeProvider => _typeProvider ?? (_typeProvider = new DefaultTypeProvider(defaultType: Report.GetType()));

        /// <summary>
        /// Instance provider. Default value is instance of <see cref="DefaultInstanceProvider"/>
        /// </summary>
        public virtual IInstanceProvider InstanceProvider => _instanceProvider ?? (_instanceProvider = new DefaultInstanceProvider(Report));

        /// <summary>
        /// Property value provider. Default value is instance of <see cref="DefaultPropertyValueProvider"/>
        /// </summary>
        public virtual IPropertyValueProvider PropertyValueProvider => _propertyValueProvider ?? (_propertyValueProvider = new DefaultPropertyValueProvider(TypeProvider, InstanceProvider));

        /// <summary>
        /// Method call value provider. Default value is instance of <see cref="DefaultMethodCallValueProvider"/>
        /// </summary>
        public virtual IMethodCallValueProvider MethodCallValueProvider => _methodCallValueProvider ?? (_methodCallValueProvider = new DefaultMethodCallValueProvider(TypeProvider, InstanceProvider));

        /// <summary>
        /// Data item value provider. Default value is instance of <see cref="DefaultDataItemValueProvider"/>
        /// </summary>
        public virtual IGenericDataItemValueProvider<HierarchicalDataItem> DataItemValueProvider => _dataItemValueProvider ?? (_dataItemValueProvider = new DefaultDataItemValueProvider());

        /// <summary>
        /// Template processor. Default value is instance of <see cref="DefaultTemplateProcessor"/>
        /// </summary>
        public virtual ITemplateProcessor TemplateProcessor => _templateProcessor ?? (_templateProcessor = new DefaultTemplateProcessor(PropertyValueProvider, SystemVariableProvider, MethodCallValueProvider, DataItemValueProvider));

        /// <summary>
        /// See <see cref="Rendering.PanelParsingSettings"/>
        /// </summary>
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

        /// <summary>
        /// Render report based on <paramref name="reportTemplate"/> and report object passed in the constructor
        /// </summary>
        /// <param name="reportTemplate">Excel report template</param>
        /// <param name="worksheets">List of worksheets that must be rendered</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="reportTemplate"/> is null</exception>
        /// <exception cref="InvalidOperationException"></exception>
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

            InitTemplateProcessor();
            InitDataItemValueProvider();

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
                IExcelPanel rootPanel = new ExcelPanel(GetRootRange(ws, worksheetPanels), Report, TemplateProcessor);
                MakePanelsHierarchy(panelsFlatView, rootPanel);
                rootPanel.Render();

                OnAfterWorksheetRender(new WorksheetRenderEventArgs { Worksheet = ws });
            }

            return reportTemplate;
        }

        private void InitTemplateProcessor()
        {
            if (TemplateProcessor is DefaultTemplateProcessor tp && SystemFunctionsType != null)
            {
                tp.SystemFunctionsType = SystemFunctionsType;
            }
        }

        private void InitDataItemValueProvider()
        {
            if (DataItemValueProvider is DefaultDataItemValueProvider p)
            {
                p.DataItemSelfTemplate = TemplateProcessor.DataItemSelfTemplate;
            }
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
            var panelFactory = new ExcelPanelFactory(Report, TemplateProcessor, PanelParsingSettings);
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
                    CheckParentChildPanelsCorrectness(parentPanel, panel);
                    parentPanel.Children.Add(panel);
                    panel.Parent = parentPanel;
                }
                else
                {
                    throw new InvalidOperationException($"Cannot find parent panel with name \"{parentPanelName}\" for panel \"{panelFlat.Key}\"");
                }
            }
        }

        private void CheckParentChildPanelsCorrectness(IExcelPanel parentPanel, IExcelPanel childPanel)
        {
            if (!ExcelHelper.IsRangeInsideAnotherRange(parentPanel.Range, childPanel.Range))
            {
                var parentNamedPanel = parentPanel as IExcelNamedPanel;
                var childNamedPanel = childPanel as IExcelNamedPanel;
                throw new InvalidOperationException(
                    $"Panel \"{parentNamedPanel?.Name ?? parentPanel.Range.ToString()}\" is not a parent of the panel \"{childNamedPanel?.Name ?? childPanel.Range.ToString()}\". Child range is outside of the parent range.");
            }
        }

        /// <summary>
        /// Raise BeforeReportRender event
        /// </summary>
        protected virtual void OnBeforeReportRender(ReportRenderEventArgs e)
        {
            BeforeReportRender?.Invoke(this, e);
        }

        /// <summary>
        /// Raise BeforeWorksheetRender event
        /// </summary>
        protected virtual void OnBeforeWorksheetRender(WorksheetRenderEventArgs e)
        {
            BeforeWorksheetRender?.Invoke(this, e);
        }

        /// <summary>
        /// Raise AfterWorksheetRender event
        /// </summary>
        protected virtual void OnAfterWorksheetRender(WorksheetRenderEventArgs e)
        {
            AfterWorksheetRender?.Invoke(this, e);
        }
    }
}