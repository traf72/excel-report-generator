using ClosedXML.Excel;
using ExcelReporter.Excel;
using ExcelReporter.Reports;

namespace ExcelReporter.Rendering.Panels.ExcelPanels
{
    internal class ExcelDataItemPanel : ExcelPanel, IDataItemPanel
    {
        public ExcelDataItemPanel(IXLRange range, IExcelReport report) : base(range, report)
        {
        }

        public HierarchicalDataItem DataItem { get; set; }

        protected override HierarchicalDataItem GetDataContext()
        {
            return DataItem;
        }

        protected override IExcelPanel CopyPanel(IXLCell cell)
        {
            IXLRange newRange = ExcelHelper.CopyRange(Range, cell);
            var panel = new ExcelDataItemPanel(newRange, Report) {DataItem = DataItem};
            FillCopyProperties(panel);
            return panel;
        }
    }
}