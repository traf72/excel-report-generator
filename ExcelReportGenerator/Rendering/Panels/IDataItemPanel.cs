namespace ExcelReportGenerator.Rendering.Panels
{
    internal interface IDataItemPanel : IPanel
    {
        HierarchicalDataItem DataItem { get; set; }
    }
}