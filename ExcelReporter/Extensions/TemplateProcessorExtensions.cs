using ExcelReporter.Rendering.TemplateProcessors;

namespace ExcelReporter.Extensions
{
    internal static class TemplateProcessorExtensions
    {
        public static string GetTemplateWithoutBorders(this ITemplateProcessor processor, string template)
        {
            template = template.Trim();
            if (processor.LeftTemplateBorder != null && template.StartsWith(processor.LeftTemplateBorder))
            {
                template = template.Substring(processor.LeftTemplateBorder.Length);
            }
            if (processor.RightTemplateBorder != null && template.EndsWith(processor.RightTemplateBorder))
            {
                template = template.Substring(0, template.Length - processor.RightTemplateBorder.Length);
            }
            return template;
        }
    }
}