using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Rendering
{
    public class HierarchicalDataItem
    {
        [LicenceKeyPart]
        public object Value { get; set; }

        public HierarchicalDataItem Parent { get; set; }
    }
}