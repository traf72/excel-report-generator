using ClosedXML.Excel;
using ExcelReportGenerator.Attributes;
using ExcelReportGenerator.Enumerators;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Excel;
using ExcelReportGenerator.Helpers;
using ExcelReportGenerator.Rendering.EventArgs;
using ExcelReportGenerator.Rendering.TemplateProcessors;

namespace ExcelReportGenerator.Rendering.Panels.ExcelPanels;

internal class ExcelDataSourcePanel : ExcelNamedPanel
{
    protected readonly string _dataSourceTemplate;
    protected readonly bool _isDataReceivedDirectly;
    protected object _data;

    private int _templatePanelRowCount;
    private int _templatePanelColumnCount;

    public ExcelDataSourcePanel(string dataSourceTemplate, IXLNamedRange namedRange, object report, ITemplateProcessor templateProcessor)
        : base(namedRange, report, templateProcessor)
    {
        if (string.IsNullOrWhiteSpace(dataSourceTemplate))
        {
            throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(dataSourceTemplate));
        }
        _dataSourceTemplate = dataSourceTemplate;
    }

    public ExcelDataSourcePanel(object data, IXLNamedRange namedRange, object report, ITemplateProcessor templateProcessor) : base(namedRange, report, templateProcessor)
    {
        _data = data ?? throw new ArgumentNullException(nameof(data), ArgumentHelper.NullParamMessage);
        _isDataReceivedDirectly = true;
    }

    [ExternalProperty]
    public string GroupBy { get; set; }

    [ExternalProperty]
    public bool GroupBlankValues { get; set; } = true;

    [ExternalProperty]
    public string BeforeDataItemRenderMethodName { get; set; }

    [ExternalProperty]
    public string AfterDataItemRenderMethodName { get; set; }

    public override void Render()
    {
        // Receive parent data item context
        HierarchicalDataItem parentDataItem = GetDataContext();

        _data = _isDataReceivedDirectly ? _data : _templateProcessor.GetValue(_dataSourceTemplate, parentDataItem);

        bool isCanceled = CallBeforeRenderMethod();
        if (isCanceled)
        {
            ResultRange = ExcelHelper.CloneRange(Range);
            return;
        }

        ICustomEnumerator enumerator = null;
        try
        {
            enumerator = EnumeratorFactory.Create(_data);
            // Removing the template if there are no data
            if (enumerator == null || enumerator.RowCount == 0)
            {
                DeletePanel(this);
                return;
            }

            // Creating the panel template which will be replicated then
            ExcelDataItemPanel templatePanel = CreateTemplatePanel();
            _templatePanelRowCount = templatePanel.Range.RowCount();
            _templatePanelColumnCount = templatePanel.Range.ColumnCount();

            // Allocating space for data
            if (enumerator.RowCount > 1)
            {
                AllocateSpaceForData(templatePanel, enumerator.RowCount);
            }

            int rowNum = 0;
            while (enumerator.MoveNext())
            {
                object currentItem = enumerator.Current;
                ExcelDataItemPanel currentPanel;
                if (rowNum != enumerator.RowCount - 1)
                {
                    IXLCell templateFirstCell = templatePanel.Range.FirstCell();
                    // The template itself is moved down or right, depending on the type of panel
                    MoveTemplatePanel(templatePanel);
                    // Copying the template on its previous place for the panel which the current data item will be rendered in
                    currentPanel = (ExcelDataItemPanel)templatePanel.Copy(templateFirstCell);
                }
                else
                {
                    // Rendering data directly in the template if there is the last data item
                    currentPanel = templatePanel;
                }

                currentPanel.DataItem = new HierarchicalDataItem { Value = currentItem, Parent = parentDataItem };

                // Fill template with data
                currentPanel.Render();
                ResultRange = ExcelHelper.MergeRanges(ResultRange, currentPanel.ResultRange);

                RemoveAllNamesRecursive(currentPanel);
                rowNum++;
            }

            RemoveName();
        }
        finally
        {
            (enumerator as IDisposable)?.Dispose();
        }

        GroupResult();
        CallAfterRenderMethod();
    }

    private void GroupResult()
    {
        if (string.IsNullOrWhiteSpace(GroupBy))
        {
            return;
        }

        int[] groupColOrRowNumbers = GroupBy.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(strNum =>
        {
            if (int.TryParse(strNum, out int num))
            {
                return num;
            }
            throw new InvalidCastException($"Parse \"{nameof(GroupBy)}\" property failed. Cannot convert value \"{strNum.Trim()}\" to {nameof(Int32)}");
        }).ToArray();

        if (Type == PanelType.Vertical)
        {
            GroupCellsVertical(ResultRange, groupColOrRowNumbers);
        }
        else
        {
            GroupCellsHorizontal(ResultRange, groupColOrRowNumbers);
        }
    }

    private void GroupCellsVertical(IXLRange range, int[] groupColNumbers)
    {
        IDictionary<int, (XLCellValue StartCellValue, int StartRowNum)> previousCellValues = new Dictionary<int, (XLCellValue, int)>();
        int rowsCount = range.Rows().Count();
        for (int rowNum = 1; rowNum <= rowsCount; rowNum++)
        {
            IXLRangeRow row = range.Row(rowNum);
            foreach (int colNum in groupColNumbers)
            {
                IXLCell currentCell = row.Cell(colNum);
                XLCellValue cellValue = currentCell.Value;
                if (previousCellValues.TryGetValue(colNum, out var previousResult))
                {
                    var cellMergedRange = new Lazy<IXLRange>(() => currentCell.MergedRange());
                    if (!previousResult.StartCellValue.Equals(cellValue)
                        && (!cellValue.IsBlank || cellMergedRange.Value == null || !cellMergedRange.Value.Contains(range.Cell(rowNum - 1, colNum))))
                    {
                        MergeCellsWithSameValue(rowNum - 1);
                        previousCellValues[colNum] = (cellValue, rowNum);
                    }
                    else if (rowNum == rowsCount)
                    {
                        MergeCellsWithSameValue(rowNum);
                    }
                }
                else
                {
                    previousCellValues[colNum] = (cellValue, rowNum);
                }
                
                void MergeCellsWithSameValue(int rowNumber)
                {
                    if (!previousResult.StartCellValue.IsBlank || GroupBlankValues)
                    {
                        range.Range(previousResult.StartRowNum, colNum, rowNumber, colNum).Merge();
                    }
                }
            }
        }
    }

    private void GroupCellsHorizontal(IXLRange range, int[] groupRowNumbers)
    {
        IDictionary<int, (XLCellValue StartCellValue, int StartColNum)> previousCellValues = new Dictionary<int, (XLCellValue, int)>();
        int colsCount = range.Columns().Count();
        for (int colNum = 1; colNum <= colsCount; colNum++)
        {
            IXLRangeColumn col = range.Column(colNum);
            foreach (int rowNum in groupRowNumbers)
            {
                IXLCell currentCell = col.Cell(rowNum);
                XLCellValue cellValue = currentCell.Value;
                if (previousCellValues.TryGetValue(rowNum, out var previousResult))
                {
                    var cellMergedRange = new Lazy<IXLRange>(() => currentCell.MergedRange());
                    if (!previousResult.StartCellValue.Equals(cellValue)
                        && (!cellValue.IsBlank || cellMergedRange.Value == null || !cellMergedRange.Value.Contains(range.Cell(rowNum, colNum - 1))))
                    {
                        MergeCellsWithSameValue(colNum - 1);
                        previousCellValues[rowNum] = (cellValue, colNum);
                    }
                    else if (colNum == colsCount)
                    {
                        MergeCellsWithSameValue(colNum);
                    }
                }
                else
                {
                    previousCellValues[rowNum] = (cellValue, colNum);
                }

                void MergeCellsWithSameValue(int columnNum)
                {
                    if (!previousResult.StartCellValue.IsBlank || GroupBlankValues)
                    {
                        range.Range(rowNum, previousResult.StartColNum, rowNum, columnNum).Merge();
                    }
                }
            }
        }
    }

    private ExcelDataItemPanel CreateTemplatePanel()
    {
        var templatePanel = new ExcelDataItemPanel(Range, _report, _templateProcessor)
        {
            Parent = Parent,
            Children = Children,
            RenderPriority = RenderPriority,
            ShiftType = ShiftType,
            Type = Type,
            BeforeRenderMethodName = BeforeDataItemRenderMethodName,
            AfterRenderMethodName = AfterDataItemRenderMethodName,
        };

        foreach (IExcelPanel child in templatePanel.Children)
        {
            child.Parent = templatePanel;
        }

        return templatePanel;
    }

    private void AllocateSpaceForData(IExcelPanel templatePanel, int dataItemsCount)
    {
        if (ShiftType == ShiftType.NoShift)
        {
            return;
        }

        IXLRange range = templatePanel.Range;
        if (Type == PanelType.Vertical)
        {
            int rowCount = (dataItemsCount - 1) * Range.RowCount();
            if (ShiftType == ShiftType.Row)
            {
                range.Worksheet.Row(range.LastRow().RowNumber()).InsertRowsBelow(rowCount);
            }
            else
            {
                range.InsertRowsBelow(rowCount, false);
            }
        }
        else
        {
            int columnCount = (dataItemsCount - 1) * Range.ColumnCount();
            if (ShiftType == ShiftType.Row)
            {
                range.Worksheet.Column(range.LastColumn().ColumnNumber()).InsertColumnsAfter(columnCount);
            }
            else
            {
                range.InsertColumnsAfter(columnCount, false);
            }
        }
    }

    private void MoveTemplatePanel(IExcelPanel templatePanel)
    {
        AddressShift shift = Type == PanelType.Vertical
            ? new AddressShift(_templatePanelRowCount, 0)
            : new AddressShift(0, _templatePanelColumnCount);

        templatePanel.Move(ExcelHelper.ShiftCell(templatePanel.Range.FirstCell(), shift));
    }

    protected void DeletePanel(IExcelPanel panel)
    {
        RemoveAllNamesRecursive(panel);
        panel.Delete();
    }

    protected override PanelBeforeRenderEventArgs GetBeforePanelRenderEventArgs()
    {
        return new DataSourcePanelBeforeRenderEventArgs { Range = Range, Data = _data };
    }

    protected override PanelEventArgs GetAfterPanelRenderEventArgs()
    {
        return new DataSourcePanelEventArgs { Range = ResultRange, Data = _data };
    }

    protected override IExcelPanel CopyPanel(IXLCell cell)
    {
        var panel = _isDataReceivedDirectly
            ? new ExcelDataSourcePanel(_data, CopyNamedRange(cell), _report, _templateProcessor)
            : new ExcelDataSourcePanel(_dataSourceTemplate, CopyNamedRange(cell), _report, _templateProcessor);

        FillCopyProperties(panel);
        return panel;
    }

    protected override void FillCopyProperties(IExcelPanel panel)
    {
        var dataSourcePanel = panel as ExcelDataSourcePanel;
        dataSourcePanel.GroupBy = GroupBy;
        dataSourcePanel.BeforeDataItemRenderMethodName = BeforeDataItemRenderMethodName;
        dataSourcePanel.AfterDataItemRenderMethodName = AfterDataItemRenderMethodName;

        base.FillCopyProperties(panel);
    }
}