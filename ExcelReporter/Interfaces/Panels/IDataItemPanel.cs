namespace ExcelReporter.Interfaces.Panels
{
    internal interface IDataItemPanel : IPanel
    {
        HierarchicalDataItem DataItem { get; set; }
    }
}