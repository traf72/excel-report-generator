namespace ExcelReporter.Rendering.TemplateProcessors
{
    /// <summary>
    /// Handles report templates
    /// </summary>
    public interface ITemplateProcessor
    {
        string TemplatePattern { get; }

        string LeftTemplateBorder { get; }

        string RightTemplateBorder { get; }

        /// <summary>
        /// Get value based on template
        /// </summary>
        /// <param name="dataItem">Data item that will be used if template is data item template</param>
        object GetValue(string template, object dataItem = null);
    }
}