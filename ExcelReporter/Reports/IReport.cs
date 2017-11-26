using ExcelReporter.Rendering.TemplateProcessors;

namespace ExcelReporter.Reports
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