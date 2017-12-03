using System;

namespace ExcelReporter.Attributes
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
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