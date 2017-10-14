namespace ExcelReporter.Interfaces.TemplateProcessors
{
    public interface ITemplateProcessor
    {
        object GetValue(string template, object dataContext = null);
    }
}