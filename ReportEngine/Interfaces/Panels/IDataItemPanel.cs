namespace ReportEngine.Interfaces.Panels
{
    public interface IDataItemPanel : IPanel
    {
        object DataItem { get; set; }
    }
}