using ExcelReporter.Rendering.TemplateProcessors;

namespace ExcelReporter.Reports
{
    public interface IReport
    {
        /// <summary>
        /// Handles report templates
        /// </summary>
        ITemplateProcessor TemplateProcessor { get; set; }

        void Run();
    }
}