using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Rendering.EventArgs;

namespace ExcelReportGenerator.Samples.Reports
{
    public class DynamicPanelDataTableConnect : DynamicPanelEnumerableConnect
    {
        public override string ReportName
        {
            get
            {
                return "Connect to DataTable via dynamic panel";
            }
        }

        public void BeforeHeadersRender(DataSourceDynamicPanelBeforeRenderEventArgs args)
        {
            args.Columns[6].AggregateFunction = AggregateFunction.Max;
            args.Columns[7].AggregateFunction = AggregateFunction.Min;
            args.Columns[8].DisplayFormat = "$#,0.00";
        }
    }
}