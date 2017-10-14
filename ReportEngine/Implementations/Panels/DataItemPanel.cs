using ClosedXML.Excel;
using JetBrains.Annotations;
using ReportEngine.Interfaces.Panels;
using ReportEngine.Interfaces.Reports;

namespace ReportEngine.Implementations.Panels
{
    public class DataItemPanel : Panel, IDataItemPanel
    {
        public DataItemPanel([NotNull] IXLRange range, [NotNull] IExcelReport report) : base(range, report)
        {
        }

        public DataItemPanel([NotNull] IExcelReport report) : base(report)
        {
        }

        public object DataItem { get; set; }

        protected override object GetDataContext()
        {
            return DataItem;
        }
    }
}