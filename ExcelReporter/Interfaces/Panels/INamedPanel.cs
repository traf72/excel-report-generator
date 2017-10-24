namespace ExcelReporter.Interfaces.Panels
{
    internal interface INamedPanel : IPanel
    {
        string Name { get; }
    }
}