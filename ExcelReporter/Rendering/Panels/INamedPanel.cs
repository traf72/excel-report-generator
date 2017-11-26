namespace ExcelReporter.Rendering.Panels
{
    internal interface INamedPanel : IPanel
    {
        string Name { get; }
    }
}