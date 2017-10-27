namespace ExcelReporter.Interfaces.TemplateProcessors
{
    public interface ITemplateProcessor
    {
        string TemplatePattern { get; }

        string LeftTemplateBorder { get; }

        string RightTemplateBorder { get; }

        object GetValue(string template, HierarchicalDataItem dataItem = null);
    }
}