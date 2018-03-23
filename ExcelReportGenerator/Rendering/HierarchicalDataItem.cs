namespace ExcelReportGenerator.Rendering
{
    /// <summary>
    /// Type that represent hierarchical data item
    /// </summary>
    public class HierarchicalDataItem
    {
        public object Value { get; set; }

        public HierarchicalDataItem Parent { get; set; }
    }
}