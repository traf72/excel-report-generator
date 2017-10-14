namespace ReportEngine.Interfaces.Panels
{
    public interface IDataSourcePanel : INamedPanel
    {
        object Data { get; }
    }
}