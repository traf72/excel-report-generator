using System;

namespace ExcelReportGenerator.Attributes
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    [LicenceKeyPart(L = true, R = true)]
    public class NullValueAttribute : Attribute
    {
        public NullValueAttribute(object value)
        {
            Value = value;
        }

        public object Value { get; }
    }
}