namespace ExcelReporter.Interfaces.TemplateProcessors
{
    public interface ITemplateProcessor
    {
        string TemplatePattern { get; }

        object GetValue(string template, HierarchicalDataItem dataItem = null);
    }
}