namespace ExcelReporter.Interfaces.Panels
{
    public interface INamedPanel : IPanel
    {
        string Name { get; }
    }
}