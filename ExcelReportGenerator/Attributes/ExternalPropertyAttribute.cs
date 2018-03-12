using System;

namespace ExcelReportGenerator.Attributes
{
    /// <summary>
    /// Mark panel property which can be populated from excel
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    internal class ExternalPropertyAttribute : Attribute
    {
        [System.Reflection.Obfuscation(Exclude = true, Feature = "renaming")]
        public Type Converter { get; set; }
    }
}