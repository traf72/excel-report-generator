using ClosedXML.Excel;
using ReportEngine.Interfaces.Panels;
using ReportEngine.Interfaces.TemplateProcessors;
using System;
using System.Linq;
using JetBrains.Annotations;
using ReportEngine.Excel;
using ReportEngine.Interfaces.Reports;

namespace ReportEngine.Implementations.Panels
{
    public class NamedPanel : Panel, INamedPanel
    {
        protected IXLNamedRange _namedRange;

        private string _copiedPanelName;

        public NamedPanel([NotNull] IXLNamedRange namedRange, [NotNull] IExcelReport report) : base(report)
        {
            if (namedRange == null)
            {
                throw new ArgumentNullException(nameof(namedRange), Constants.NullParamMessage);
            }

            _namedRange = namedRange;
        }

        public virtual string Name => _namedRange.Name;

        public override IXLRange Range => _namedRange.Ranges.ElementAt(0);

        public override IPanel Copy(IXLCell cell, bool recursive = true)
        {
            string name = $"{Name}_{Guid.NewGuid():N}";
            return Copy(cell, name, recursive);
        }

        public INamedPanel Copy(IXLCell cell, string name, bool recursive = true)
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
            return (INamedPanel)base.Copy(cell, recursive);
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

        protected override IPanel CopyPanel(IXLCell cell)
        {
            var panel = new NamedPanel(CopyNamedRange(cell), Report);
            FillCopyProperties(panel);
            return panel;
        }

        protected IXLNamedRange CopyNamedRange(IXLCell cell)
        {
            return ExcelHelper.CopyNamedRange(_namedRange, cell, _copiedPanelName);
        }

        protected override IPanel CopyChild(IPanel fromChild, IXLCell cell)
        {
            INamedPanel namedChild = fromChild as INamedPanel;
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

        protected INamedPanel GetNearestNamedParent()
        {
            IPanel parent = Parent;
            while (parent != null)
            {
                INamedPanel namedParent = parent as INamedPanel;
                if (namedParent != null)
                {
                    return namedParent;
                }
                parent = parent.Parent;
            }
            return null;
        }

        public static void RemoveAllNamesRecursive(IPanel panel)
        {
            foreach (IPanel p in panel.Children)
            {
                RemoveAllNamesRecursive(p);
            }

            INamedPanel namedChild = panel as INamedPanel;
            namedChild?.RemoveName();
        }
    }
}