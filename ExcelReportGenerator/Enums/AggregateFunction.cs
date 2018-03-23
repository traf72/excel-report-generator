using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Enums
{
    /// <summary>
    /// Built-in aggregate functions
    /// </summary>
    [LicenceKeyPart(L = true)]
    public enum AggregateFunction
    {
        Sum,
        Count,
        Avg,
        Max,
        Min,
        Custom,
        NoAggregation,
    }
}