namespace ExcelReporter.Interfaces.TemplateProcessors
{
    public interface IGenericTemplateProcessor<in T> : ITemplateProcessor
    {
        object GetValue(string template, T dataItem = default(T));
    }
}