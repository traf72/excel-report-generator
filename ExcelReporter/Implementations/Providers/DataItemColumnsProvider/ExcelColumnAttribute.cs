using System;

namespace ExcelReporter.Implementations.Providers.DataItemColumnsProvider
{
    // Пока оставил возможность применения только к свойствам
    //[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnAttribute : Attribute
    {
        /// <summary>
        /// Column caption which will be shown in excel
        /// </summary>
        public string Caption { get; set; }

        /// <summary>
        /// Column width
        /// </summary>
        public double Width { get; set; }
    }
}