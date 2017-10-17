namespace ExcelReporter.Interfaces.TemplateProcessors
{
    public interface ITemplateProcessor
    {
        string Pattern { get; }

        object GetValue(string template, HierarchicalDataItem dataItem = null);
    }
}