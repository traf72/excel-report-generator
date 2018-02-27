using ClosedXML.Excel;
using ExcelReportGenerator.Excel;
using ExcelReportGenerator.Rendering.EventArgs;
using ExcelReportGenerator.Rendering.TemplateProcessors;

namespace ExcelReportGenerator.Rendering.Panels.ExcelPanels
{
    internal class ExcelDataItemPanel : ExcelPanel, IDataItemPanel
    {
        public ExcelDataItemPanel(IXLRange range, object report, ITemplateProcessor templateProcessor)
            : base(range, report, templateProcessor)
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
            return new DataItemPanelEventArgs { Range = ResultRange, DataItem = DataItem };
        }

        // TODO Покрыть Unit-тестами
        protected override IExcelPanel CopyPanel(IXLCell cell)
        {
            IXLRange newRange = ExcelHelper.CopyRange(Range, cell);
            var panel = new ExcelDataItemPanel(newRange, _report, _templateProcessor) { DataItem = DataItem };
            FillCopyProperties(panel);
            return panel;
        }
    }
}