using System;

namespace ExcelReportGenerator.Attributes
{
    [AttributeUsage(AttributeTargets.Interface | AttributeTargets.Class | AttributeTargets.Struct | AttributeTargets.Enum | AttributeTargets.Method | AttributeTargets.Property)]
    internal class LicenceKeyPartAttribute : Attribute
    {
        public LicenceKeyPartAttribute()
        {
        }

        public bool L { get; set; }

        public bool R { get; set; }

        public bool U { get; set; }
    }
}