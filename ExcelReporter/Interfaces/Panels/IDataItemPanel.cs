namespace ExcelReporter.Interfaces.Panels
{
    public interface IDataItemPanel : IPanel
    {
        object DataItem { get; set; }
    }
}