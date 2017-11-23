using System;

namespace ExcelReporter.Implementations.Providers.DataItemColumnsProvider
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    public class ExcelColumnAttribute : Attribute
    {
        public string Caption { get; set; }

        public double? Width { get; set; }
    }
}