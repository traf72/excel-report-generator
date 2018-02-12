namespace ExcelReportGenerator.Rendering.TemplateProcessors
{
    /// <summary>
    /// Handles report templates
    /// </summary>
    public interface ITemplateProcessor
    {
        /// <summary>
        /// Left template border
        /// Cannot be null, empty or whitespace
        /// </summary>
        string LeftTemplateBorder { get; }

        /// <summary>
        /// Right template border
        /// Cannot be null, empty or whitespace
        /// </summary>
        string RightTemplateBorder { get; }

        /// <summary>
        /// Separator between member label (Property, MethodCall, DataItem) and member template
        /// Cannot be null or empty
        /// </summary>
        string MemberLabelSeparator { get; }

        /// <summary>
        /// Property label
        /// Cannot be null, empty or whitespace
        /// </summary>
        string PropertyMemberLabel { get; }

        /// <summary>
        /// Method call label
        /// Cannot be null, empty or whitespace
        /// </summary>
        string MethodCallMemberLabel { get; }

        /// <summary>
        /// Data item label
        /// Cannot be null, empty or whitespace
        /// </summary>
        string DataItemMemberLabel { get; }

        /// <summary>
        /// System variable label
        /// Cannot be null, empty or whitespace
        /// </summary>
        string SystemVariableMemberLabel { get; }

        /// <summary>
        /// System function label
        /// Cannot be null, empty or whitespace
        /// </summary>
        string SystemFunctionMemberLabel { get; }

        /// <summary>
        /// Horizontal page break label
        /// Cannot be null, empty or whitespace
        /// </summary>
        string HorizontalPageBreakLabel { get; }

        /// <summary>
        /// Vertical page break label
        /// Cannot be null, empty or whitespace
        /// </summary>
        string VerticalPageBreakLabel { get; }

        /// <summary>
        /// Get value based on template
        /// </summary>
        /// <param name="dataItem">Data item that will be used if template is data item template</param>
        object GetValue(string template, HierarchicalDataItem dataItem = null);
    }
}