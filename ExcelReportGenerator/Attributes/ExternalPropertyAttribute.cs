namespace ExcelReportGenerator.Attributes;

/// <summary>
/// Marks panel property which can be populated from excel
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
internal class ExternalPropertyAttribute : Attribute
{
    public Type Converter { get; set; }
}