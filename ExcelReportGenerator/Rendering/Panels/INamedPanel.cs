namespace ExcelReportGenerator.Rendering.Panels
{
    internal interface INamedPanel : IPanel
    {
        string Name { get; }
    }
}