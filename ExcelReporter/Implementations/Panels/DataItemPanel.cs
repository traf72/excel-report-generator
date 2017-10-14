using ClosedXML.Excel;
using ExcelReporter.Interfaces.Panels;
using ExcelReporter.Interfaces.Reports;

namespace ExcelReporter.Implementations.Panels
{
    public class DataItemPanel : Panel, IDataItemPanel
    {
        public DataItemPanel(IXLRange range, IExcelReport report) : base(range, report)
        {
        }

        public DataItemPanel(IExcelReport report) : base(report)
        {
        }

        public object DataItem { get; set; }

        protected override object GetDataContext()
        {
            return DataItem;
        }
    }
}