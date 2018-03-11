using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Enums
{
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