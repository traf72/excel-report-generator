using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using ExcelReporter.Enums;
using ExcelReporter.Excel;
using ExcelReporter.Interfaces.Panels;
using ExcelReporter.Interfaces.Panels.Excel;
using ExcelReporter.Interfaces.Reports;

namespace ExcelReporter.Implementations.Panels.Excel
{
    public class ExcelPanel : IExcelPanel
    {
        protected IExcelPanel _parent;

        protected RangeCoords _coordsRelativeParent;

        public ExcelPanel(IXLRange range, IExcelReport report)
        {
            if (range == null)
            {
                throw new ArgumentNullException(nameof(range), Constants.NullParamMessage);
            }
            if (report == null)
            {
                throw new ArgumentNullException(nameof(report), Constants.NullParamMessage);
            }

            Range = range;
            Report = report;
        }

        protected ExcelPanel(IExcelReport report)
        {
            if (report == null)
            {
                throw new ArgumentNullException(nameof(report), Constants.NullParamMessage);
            }

            Report = report;
        }

        public IExcelReport Report { get; set; }

        public virtual IXLRange Range { get; private set; }

        public IExcelPanel Parent
        {
            get { return _parent; }
            set
            {
                _parent = value;
                _coordsRelativeParent = _parent != null ? ExcelHelper.GetRangeCoordsRelativeParent(_parent.Range, Range) : new RangeCoords();
            }
        }

        public IEnumerable<IExcelPanel> Children { get; set; } = new List<IExcelPanel>();

        public ShiftType ShiftType { get; set; }

        public PanelType Type { get; set; }

        public int RenderPriority { get; set; }

        public virtual void Render()
        {
            IList<IXLCell> childrenCells = Children.SelectMany(c => c.Range.CellsUsed()).ToList();
            foreach (IXLCell cell in Range.CellsUsed().Where(c => !childrenCells.Contains(c)))
            {
                string cellValue = cell.Value.ToString();
                MatchCollection matches = Regex.Matches(cellValue, Report.TemplateProcessor.TemplatePattern);
                if (matches.Count == 0)
                {
                    continue;
                }
                if (matches.Count == 1 && Regex.IsMatch(cellValue, $@"^{Report.TemplateProcessor.TemplatePattern}$"))
                {
                    cell.Value = Report.TemplateProcessor.GetValue(cellValue, GetDataContext());
                    continue;
                }

                foreach (object match in matches)
                {
                    string template = match.ToString();
                    cellValue = cellValue.Replace(template, Report.TemplateProcessor.GetValue(template).ToString());
                }

                cell.Value = cellValue;
            }

            foreach (IExcelPanel child in Children.OrderByDescending(p => p.RenderPriority))
            {
                child.Render();
            }
        }

        public virtual IExcelPanel Copy(IXLCell cell, bool recursive = true)
        {
            if (cell == null)
            {
                throw new ArgumentNullException(nameof(cell), Constants.NullParamMessage);
            }

            IExcelPanel newPanel = CopyPanel(cell);
            SetParent(Parent, newPanel);
            if (!recursive)
            {
                return newPanel;
            }

            IList<IExcelPanel> children = new List<IExcelPanel>(Children.Count());
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
                throw new ArgumentNullException(nameof(cell), Constants.NullParamMessage);
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

        public virtual void RecalculateRangeRelativeParentRecursive()
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
        }

        protected virtual IExcelPanel CopyPanel(IXLCell cell)
        {
            IXLRange newRange = ExcelHelper.CopyRange(Range, cell);
            var panel = new ExcelPanel(newRange, Report);
            FillCopyProperties(panel);
            return panel;
        }

        protected virtual HierarchicalDataItem GetDataContext()
        {
            IExcelPanel parent = Parent;
            while (parent != null)
            {
                IDataItemPanel dataItemPanel = parent as IDataItemPanel;
                if (dataItemPanel != null)
                {
                    return dataItemPanel.DataItem;
                }
                parent = parent.Parent;
            }
            return null;
        }
    }
}