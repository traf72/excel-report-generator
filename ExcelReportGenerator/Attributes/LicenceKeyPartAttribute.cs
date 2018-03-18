using System;
using System.Reflection;

namespace ExcelReportGenerator.Attributes
{
    [AttributeUsage(AttributeTargets.Interface | AttributeTargets.Class | AttributeTargets.Struct | AttributeTargets.Enum | AttributeTargets.Method | AttributeTargets.Property)]
    internal class LicenceKeyPartAttribute : Attribute
    {
        // Reverse
        [Obfuscation(Exclude = true, Feature = "renaming")]
        public bool L { get; set; }

        // ToUpper
        [Obfuscation(Exclude = true, Feature = "renaming")]
        public bool R { get; set; }

        // ToLower
        [Obfuscation(Exclude = true, Feature = "renaming")]
        public bool U { get; set; }
    }
}