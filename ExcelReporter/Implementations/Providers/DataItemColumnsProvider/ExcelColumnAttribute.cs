using System;

namespace ExcelReporter.Implementations.Providers.DataItemColumnsProvider
{
    // Пока оставил возможность применения только к свойствам
    //[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnAttribute : Attribute
    {
        public string Caption { get; set; }

        public double Width { get; set; }
    }
}