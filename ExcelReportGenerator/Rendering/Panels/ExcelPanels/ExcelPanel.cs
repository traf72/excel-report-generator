using ClosedXML.Excel;
using ExcelReportGenerator.Attributes;
using ExcelReportGenerator.Converters.ExternalPropertiesConverters;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Excel;
using ExcelReportGenerator.Exceptions;
using ExcelReportGenerator.Extensions;
using ExcelReportGenerator.Helpers;
using ExcelReportGenerator.Rendering.EventArgs;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using System.Reflection;
using System.Text.RegularExpressions;

namespace ExcelReportGenerator.Rendering.Panels.ExcelPanels;

internal class ExcelPanel : IExcelPanel
{
    protected IExcelPanel _parent;

    protected RangeCoords _coordsRelativeParent;

    protected readonly object _report;

    protected readonly ITemplateProcessor _templateProcessor;

    public ExcelPanel(IXLRange range, object report, ITemplateProcessor templateProcessor)
    {
        Range = range ?? throw new ArgumentNullException(nameof(range), ArgumentHelper.NullParamMessage);
        _report = report ?? throw new ArgumentNullException(nameof(report), ArgumentHelper.NullParamMessage);
        _templateProcessor = templateProcessor ?? throw new ArgumentNullException(nameof(templateProcessor), ArgumentHelper.NullParamMessage);
    }

    protected ExcelPanel(object report, ITemplateProcessor templateProcessor)
    {
        _report = report ?? throw new ArgumentNullException(nameof(report), ArgumentHelper.NullParamMessage);
        _templateProcessor = templateProcessor ?? throw new ArgumentNullException(nameof(templateProcessor), ArgumentHelper.NullParamMessage);
    }

    public virtual IXLRange Range { get; private set; }

    public virtual IXLRange ResultRange { get; protected set; }

    public IExcelPanel Parent
    {
        get => _parent;
        set
        {
            _parent = value;
            _coordsRelativeParent = _parent != null ? ExcelHelper.GetRangeCoordsRelativeParent(_parent.Range, Range, false) : new RangeCoords();
        }
    }

    public IList<IExcelPanel> Children { get; set; } = new List<IExcelPanel>();

    [ExternalProperty(Converter = typeof(ShiftTypeConverter))]
    public ShiftType ShiftType { get; set; }

    [ExternalProperty(Converter = typeof(PanelTypeConverter))]
    public PanelType Type { get; set; }

    [ExternalProperty(Converter = typeof(RenderPriorityConverter))]
    public int RenderPriority { get; set; }

    [ExternalProperty]
    public string BeforeRenderMethodName { get; set; }

    [ExternalProperty]
    public string AfterRenderMethodName { get; set; }

    public virtual void Render()
    {
        ResultRange = ExcelHelper.CloneRange(Range);

        bool isCanceled = CallBeforeRenderMethod();
        if (isCanceled)
        {
            return;
        }

        IList<IXLCell> childrenCells = Children.SelectMany(c => c.Range.CellsUsed(XLCellsUsedOptions.Contents)).ToList();
        string templatePattern = _templateProcessor.GetFullRegexPattern();
        foreach (IXLCell cell in Range.CellsUsedWithoutFormulas(c => !childrenCells.Contains(c)))
        {
            string cellValue = cell.Value.ToString();
            if (_templateProcessor.IsHorizontalPageBreak(cellValue))
            {
                cell.Worksheet.PageSetup.AddHorizontalPageBreak(cell.WorksheetRow().RowNumber());
                cell.Value = Blank.Value;
                continue;
            }
            if (_templateProcessor.IsVerticalPageBreak(cellValue))
            {
                cell.Worksheet.PageSetup.AddVerticalPageBreak(cell.WorksheetColumn().ColumnNumber());
                cell.Value = Blank.Value;
                continue;
            }

            MatchCollection matches = GetMatches(cellValue);
            if (matches.Count == 0)
            {
                continue;
            }

            if (!cell.HasRichText)
            {
                cell.Value = GetCellValue(cellValue);
            }
            else
            {
                foreach (IXLRichString richString in cell.GetRichText())
                {
                    matches = GetMatches(richString.Text);
                    if (matches.Count == 0)
                    {
                        continue;
                    }

                    richString.Text = GetCellValue(richString.Text).ToString();
                }
            }

            continue;
            
            MatchCollection GetMatches(string cellVal) => Regex.Matches(cellVal, templatePattern, RegexOptions.IgnoreCase);
            
            XLCellValue GetCellValue(string cellVal)
            {
                HierarchicalDataItem dataContext = GetDataContext();
                if (matches.Count == 1 && Regex.IsMatch(cellVal, $@"^{templatePattern}$", RegexOptions.IgnoreCase))
                {
                    object value = _templateProcessor.GetValue(cellVal, dataContext);
                    if (value == null)
                    {
                        return Blank.Value;
                    }

                    if (value.GetType().IsNumeric())
                    {
                        return Convert.ToDouble(value);
                    }

                    return value switch
                    {
                        bool boolValue => boolValue,
                        DateTime dateTimeValue => dateTimeValue,
                        TimeSpan timeSpanValue => timeSpanValue,
                        _ => GetValueFromString(value.ToString())
                    };
                }

                foreach (object match in matches)
                {
                    string template = match.ToString();
                    cellVal = cellVal.Replace(template,
                        _templateProcessor.GetValue(template, dataContext)?.ToString());
                }

                return GetValueFromString(cellVal);
            }

            XLCellValue GetValueFromString(string value) => value == string.Empty ? Blank.Value : value;
        }

        foreach (IExcelPanel child in Children.OrderByDescending(p => p.RenderPriority))
        {
            child.Render();
            ResultRange = ExcelHelper.MergeRanges(ResultRange, child.ResultRange);
        }

        CallAfterRenderMethod();
    }

    public virtual IExcelPanel Copy(IXLCell cell, bool recursive = true)
    {
        if (cell == null)
        {
            throw new ArgumentNullException(nameof(cell), ArgumentHelper.NullParamMessage);
        }

        IExcelPanel newPanel = CopyPanel(cell);
        SetParent(Parent, newPanel);
        if (!recursive)
        {
            return newPanel;
        }

        IList<IExcelPanel> children = new List<IExcelPanel>(Children.Count);
        foreach (IExcelPanel child in Children)
        {
            CellCoords firstCellRelativeCoords = ExcelHelper.GetCellCoordsRelativeRange(Range, child.Range.FirstCell());
            IXLCell firstCell = newPanel.Range.Cell(firstCellRelativeCoords.RowNum, firstCellRelativeCoords.ColNum);
            IExcelPanel newChild = CopyChild(child, firstCell);
            SetParent(newPanel, newChild);
            children.Add(newChild);
        }

        newPanel.Children = children;
        return newPanel;
    }

    public virtual void Move(IXLCell cell)
    {
        if (cell == null)
        {
            throw new ArgumentNullException(nameof(cell), ArgumentHelper.NullParamMessage);
        }

        MoveRange(cell);
        SetParent(Parent, this);
        MoveChildren();
    }

    protected virtual void MoveRange(IXLCell cell)
    {
        Range = ExcelHelper.MoveRange(Range, cell);
    }

    protected virtual void MoveChildren()
    {
        foreach (IExcelPanel child in Children)
        {
            child.RecalculateRangeRelativeParentRecursive();
        }
    }

    void IExcelPanel.RecalculateRangeRelativeParentRecursive()
    {
        if (Parent != null)
        {
            Range = Parent.Range.Range(
                _coordsRelativeParent.FirstCell.RowNum,
                _coordsRelativeParent.FirstCell.ColNum,
                _coordsRelativeParent.LastCell.RowNum,
                _coordsRelativeParent.LastCell.ColNum);
            MoveChildren();
        }
    }

    protected virtual IExcelPanel CopyChild(IExcelPanel fromChild, IXLCell cell)
    {
        return fromChild.Copy(cell);
    }

    public virtual void Delete()
    {
        ExcelHelper.DeleteRange(Range, ShiftType, Type == PanelType.Vertical ? XLShiftDeletedCells.ShiftCellsUp : XLShiftDeletedCells.ShiftCellsLeft);
    }

    protected void SetParent(IExcelPanel probableParent, IExcelPanel child)
    {
        if (probableParent != null && ExcelHelper.IsRangeInsideAnotherRange(probableParent.Range, child.Range))
        {
            child.Parent = probableParent;
        }
        else
        {
            child.Parent = null;
        }
    }

    protected virtual void FillCopyProperties(IExcelPanel panel)
    {
        panel.Type = Type;
        panel.ShiftType = ShiftType;
        panel.RenderPriority = RenderPriority;
        panel.BeforeRenderMethodName = BeforeRenderMethodName;
        panel.AfterRenderMethodName = AfterRenderMethodName;
    }

    protected virtual IExcelPanel CopyPanel(IXLCell cell)
    {
        var panel = new ExcelPanel(CopyRange(cell), _report, _templateProcessor);
        FillCopyProperties(panel);
        return panel;
    }

    protected IXLRange CopyRange(IXLCell cell)
    {
        return ExcelHelper.CopyRange(Range, cell);
    }

    protected virtual HierarchicalDataItem GetDataContext()
    {
        IExcelPanel parent = Parent;
        while (parent != null)
        {
            if (parent is IDataItemPanel dataItemPanel)
            {
                return dataItemPanel.DataItem;
            }
            parent = parent.Parent;
        }
        return null;
    }

    protected bool CallBeforeRenderMethod()
    {
        if (!string.IsNullOrWhiteSpace(BeforeRenderMethodName))
        {
            PanelBeforeRenderEventArgs eventArgs = GetBeforePanelRenderEventArgs();
            CallReportMethod(BeforeRenderMethodName, new object[] { eventArgs });
            return eventArgs.IsCanceled;
        }
        return false;
    }

    protected virtual PanelBeforeRenderEventArgs GetBeforePanelRenderEventArgs()
    {
        return new PanelBeforeRenderEventArgs { Range = Range };
    }

    protected void CallAfterRenderMethod()
    {
        if (!string.IsNullOrWhiteSpace(AfterRenderMethodName))
        {
            CallReportMethod(AfterRenderMethodName, new object[] { GetAfterPanelRenderEventArgs() });
        }
    }

    protected virtual PanelEventArgs GetAfterPanelRenderEventArgs()
    {
        return new PanelEventArgs { Range = ResultRange };
    }

    protected object CallReportMethod(string methodName, object[] parameters = null)
    {
        if (string.IsNullOrWhiteSpace(methodName))
        {
            throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(methodName));
        }

        MethodInfo method = _report.GetType().GetMethod(methodName, BindingFlags.Instance | BindingFlags.Public);
        if (method == null)
        {
            throw new MethodNotFoundException($"Cannot find public instance method \"{methodName}\" in type \"{_report.GetType().Name}\"");
        }
        return method.Invoke(_report, parameters);
    }
}