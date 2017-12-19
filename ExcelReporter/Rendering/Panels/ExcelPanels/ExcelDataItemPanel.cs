using ClosedXML.Excel;
using ExcelReporter.Excel;
using ExcelReporter.Rendering.EventArgs;
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

        protected override PanelBeforeRenderEventArgs GetBeforePanelRenderEventArgs()
        {
            return new DataItemPanelBeforeRenderEventArgs { Range = Range, DataItem = DataItem };
        }

        protected override PanelEventArgs GetAfterPanelRenderEventArgs()
        {
            return new DataItemPanelEventArgs { Range = Range, DataItem = DataItem };
        }

        // TODO Покрыть Unit-тестами
        protected override IExcelPanel CopyPanel(IXLCell cell)
        {
            IXLRange newRange = ExcelHelper.CopyRange(Range, cell);
            var panel = new ExcelDataItemPanel(newRange, Report) { DataItem = DataItem };
            FillCopyProperties(panel);
            return panel;
        }
    }
}