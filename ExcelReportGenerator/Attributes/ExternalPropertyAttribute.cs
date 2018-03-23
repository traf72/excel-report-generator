using System;

namespace ExcelReportGenerator.Attributes
{
    // Marks panel property which can be populated from excel
    [AttributeUsage(AttributeTargets.Property)]
    internal class ExternalPropertyAttribute : Attribute
    {
        [System.Reflection.Obfuscation(Exclude = true, Feature = "renaming")]
        public Type Converter { get; set; }
    }
}