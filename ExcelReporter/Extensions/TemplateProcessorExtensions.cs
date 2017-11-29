using ExcelReporter.Rendering.TemplateProcessors;

namespace ExcelReporter.Extensions
{
    internal static class TemplateProcessorExtensions
    {
        /// <summary>
        /// Remove template borders
        /// </summary>
        public static string UnwrapTemplate(this ITemplateProcessor processor, string template)
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

        /// <summary>
        /// Wrap template with borders
        /// </summary>
        public static string WrapTemplate(this ITemplateProcessor processor, string template)
        {
            return $"{processor.LeftTemplateBorder}{template}{processor.RightTemplateBorder}";
        }
    }
}