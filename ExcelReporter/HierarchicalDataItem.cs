namespace ExcelReporter
{
    public class HierarchicalDataItem
    {
        public object Value { get; set; }

        public HierarchicalDataItem Parent { get; set; }
    }
}