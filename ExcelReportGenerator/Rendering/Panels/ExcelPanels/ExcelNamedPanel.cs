using ClosedXML.Excel;
using ExcelReportGenerator.Excel;
using ExcelReportGenerator.Helpers;
using ExcelReportGenerator.Rendering.TemplateProcessors;

namespace ExcelReportGenerator.Rendering.Panels.ExcelPanels;

internal class ExcelNamedPanel : ExcelPanel, IExcelNamedPanel
{
    protected IXLDefinedName _namedRange;

    private string _copiedPanelName;

    public ExcelNamedPanel(IXLDefinedName namedRange, object report, ITemplateProcessor templateProcessor) : base(report, templateProcessor)
    {
        _namedRange = namedRange ?? throw new ArgumentNullException(nameof(namedRange), ArgumentHelper.NullParamMessage);
    }

    public virtual string Name => _namedRange.Name;

    public override IXLRange Range => _namedRange.Ranges.ElementAt(0);

    public override void Render()
    {
        base.Render();
        RemoveName();
    }

    public override IExcelPanel Copy(IXLCell cell, bool recursive = true)
    {
        string name = $"{Name}_{Guid.NewGuid():N}";
        return Copy(cell, name, recursive);
    }

    public IExcelNamedPanel Copy(IXLCell cell, string name, bool recursive = true)
    {
        if (cell == null)
        {
            throw new ArgumentNullException(nameof(cell), ArgumentHelper.NullParamMessage);
        }
        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(name));
        }

        _copiedPanelName = name;
        return (IExcelNamedPanel)base.Copy(cell, recursive);
    }

    public void RemoveName(bool recursive = false)
    {
        if (recursive)
        {
            RemoveAllNamesRecursive(this);
        }
        else
        {
            _namedRange.Delete();
        }
    }

    public override void Delete()
    {
        RemoveName(true);
        base.Delete();
    }

    protected override IExcelPanel CopyPanel(IXLCell cell)
    {
        var panel = new ExcelNamedPanel(CopyNamedRange(cell), _report, _templateProcessor);
        FillCopyProperties(panel);
        return panel;
    }

    protected IXLDefinedName CopyNamedRange(IXLCell cell)
    {
        return ExcelHelper.CopyNamedRange(_namedRange, cell, _copiedPanelName);
    }

    protected override IExcelPanel CopyChild(IExcelPanel fromChild, IXLCell cell)
    {
        return fromChild is IExcelNamedPanel namedChild ? namedChild.Copy(cell, $"{_copiedPanelName}_{namedChild.Name}") : fromChild.Copy(cell);
    }

    protected override void MoveRange(IXLCell cell)
    {
        _namedRange = ExcelHelper.MoveNamedRange(_namedRange, cell);
    }

    void IExcelPanel.RecalculateRangeRelativeParentRecursive()
    {
        if (Parent == null)
        {
            return;
        }

        string name = _namedRange.Name;
        IXLRange range = Parent.Range.Range(
            _coordsRelativeParent.FirstCell.RowNum,
            _coordsRelativeParent.FirstCell.ColNum,
            _coordsRelativeParent.LastCell.RowNum,
            _coordsRelativeParent.LastCell.ColNum);
        _namedRange.Delete();
        range.AddToNamed(name, XLScope.Worksheet);
        _namedRange = range.Worksheet.NamedRange(name);
        MoveChildren();
    }

    protected IExcelNamedPanel GetNearestNamedParent()
    {
        IExcelPanel parent = Parent;
        while (parent != null)
        {
            if (parent is IExcelNamedPanel namedParent)
            {
                return namedParent;
            }
            parent = parent.Parent;
        }
        return null;
    }

    public static void RemoveAllNamesRecursive(IExcelPanel panel)
    {
        foreach (IExcelPanel p in panel.Children)
        {
            RemoveAllNamesRecursive(p);
        }

        IExcelNamedPanel namedChild = panel as IExcelNamedPanel;
        namedChild?.RemoveName();
    }
}