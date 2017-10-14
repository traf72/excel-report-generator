using ExcelReporter.Interfaces.TemplateProcessors;

namespace ExcelReporter.Interfaces.Reports
{
    public interface IReport
    {
        void Run();

        ITemplateProcessor TemplateProcessor { get; set; }
    }
}