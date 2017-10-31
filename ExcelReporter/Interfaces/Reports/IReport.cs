using ExcelReporter.Interfaces.TemplateProcessors;

namespace ExcelReporter.Interfaces.Reports
{
    public interface IReport
    {
        void Run();

        /// <summary>
        /// Handles report templates
        /// </summary>
        ITemplateProcessor TemplateProcessor { get; set; }
    }
}