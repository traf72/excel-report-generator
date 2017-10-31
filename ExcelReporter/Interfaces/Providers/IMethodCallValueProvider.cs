using ExcelReporter.Interfaces.TemplateProcessors;

namespace ExcelReporter.Interfaces.Providers
{
    public interface IMethodCallValueProvider
    {
        object CallMethod(string methodCallTemplate, ITemplateProcessor templateProcessor, object dataItem, bool isStatic = false);
    }
}