using ReportEngine.Interfaces.TemplateProcessors;

namespace ReportEngine.Interfaces.Reports
{
    public interface IReport
    {
        void Run();

        ITemplateProcessor TemplateProcessor { get; set; }
    }
}