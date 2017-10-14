namespace ExcelReporter.Interfaces.Panels
{
    public interface IDataSourcePanel : INamedPanel
    {
        object Data { get; }
    }
}