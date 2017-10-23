using ClosedXML.Excel;
using ExcelReporter.Excel;
using ExcelReporter.Interfaces.Panels.Excel;
using ExcelReporter.Interfaces.Reports;
using System;
using System.Linq;

namespace ExcelReporter.Implementations.Panels.Excel
{
    public class ExcelNamedPanel : ExcelPanel, IExcelNamedPanel
    {
        protected IXLNamedRange _namedRange;

        private string _copiedPanelName;

        public ExcelNamedPanel(IXLNamedRange namedRange, IExcelReport report) : base(report)
        {
            if (namedRange == null)
            {
                throw new ArgumentNullException(nameof(namedRange), Constants.NullParamMessage);
            }

            _namedRange = namedRange;
        }

        public virtual string Name => _namedRange.Name;

        public override IXLRange Range => _namedRange.Ranges.ElementAt(0);

        public override IExcelPanel Copy(IXLCell cell, bool recursive = true)
        {
            string name = $"{Name}_{Guid.NewGuid():N}";
            return Copy(cell, name, recursive);
        }

        public IExcelNamedPanel Copy(IXLCell cell, string name, bool recursive = true)
        {
            if (cell == null)
            {
                throw new ArgumentNullException(nameof(cell), Constants.NullParamMessage);
            }
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException(Constants.EmptyStringParamMessage, nameof(name));
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
            var panel = new ExcelNamedPanel(CopyNamedRange(cell), Report);
            FillCopyProperties(panel);
            return panel;
        }

        protected IXLNamedRange CopyNamedRange(IXLCell cell)
        {
            return ExcelHelper.CopyNamedRange(_namedRange, cell, _copiedPanelName);
        }

        protected override IExcelPanel CopyChild(IExcelPanel fromChild, IXLCell cell)
        {
            IExcelNamedPanel namedChild = fromChild as IExcelNamedPanel;
            return namedChild != null ? namedChild.Copy(cell, $"{_copiedPanelName}_{namedChild.Name}") : fromChild.Copy(cell);
        }

        protected override void MoveRange(IXLCell cell)
        {
            _namedRange = ExcelHelper.MoveNamedRange(_namedRange, cell);
        }

        public override void RecalculateRangeRelativeParentRecursive()
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
                IExcelNamedPanel namedParent = parent as IExcelNamedPanel;
                if (namedParent != null)
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
}