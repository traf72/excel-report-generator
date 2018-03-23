using System;

namespace ExcelReportGenerator.Attributes
{
    /// <summary>
    /// An attribute that allows you to mark property as no excel column
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    public class NoExcelColumnAttribute : Attribute
    {
    }
}