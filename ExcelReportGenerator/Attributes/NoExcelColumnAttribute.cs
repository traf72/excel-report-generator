using System;

namespace ExcelReportGenerator.Attributes
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    public class NoExcelColumnAttribute : Attribute
    {
    }
}