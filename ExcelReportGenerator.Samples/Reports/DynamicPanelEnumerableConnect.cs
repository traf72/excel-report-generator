using ClosedXML.Excel;
using ExcelReportGenerator.Rendering.EventArgs;

namespace ExcelReportGenerator.Samples.Reports
{
    public class DynamicPanelEnumerableConnect : ReportBase
    {
        public override string ReportName => "Connect to IEnumerable via dynamic panel";

        public void AfterDataTemplatesRender(DataSourceDynamicPanelEventArgs args)
        {
            args.Range.FirstCell().Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Background2);
            args.Range.Range(1, 7, 1, 9).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        }

        public void AfterTotalsRender(DataSourcePanelEventArgs args)
        {
            args.Range.Range(1, 7, 1, 9).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            IXLRange mergedRange = args.Range.Range(1, 1, 1, 6).Merge();
            mergedRange.Value = "Totals";
            mergedRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        }
    }
}