namespace ExcelReporter.Rendering.TemplateProcessors
{
    /// <summary>
    /// Handles report templates
    /// </summary>
    /// <typeparam name="T">Type of data item object</typeparam>
    public interface IGenericTemplateProcessor<in T> : ITemplateProcessor
    {
        /// <summary>
        /// Get value based on template
        /// </summary>
        /// <param name="dataItem">Data item that will be used if template is data item template</param>
        object GetValue(string template, T dataItem = default);
    }
}