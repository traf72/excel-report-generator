using ClosedXML.Excel;
using ExcelReporter.Excel;
using ExcelReporter.Interfaces.Panels;
using ExcelReporter.Interfaces.Reports;

namespace ExcelReporter.Implementations.Panels
{
    internal class DataItemPanel : Panel, IDataItemPanel
    {
        public DataItemPanel(IXLRange range, IExcelReport report) : base(range, report)
        {
        }

        public HierarchicalDataItem DataItem { get; set; }

        protected override HierarchicalDataItem GetDataContext()
        {
            return DataItem;
        }

        protected override IPanel CopyPanel(IXLCell cell)
        {
            IXLRange newRange = ExcelHelper.CopyRange(Range, cell);
            var panel = new DataItemPanel(newRange, Report) {DataItem = DataItem};
            FillCopyProperties(panel);
            return panel;
        }
    }
}