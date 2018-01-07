using System;

namespace ExcelReportGenerator.Attributes
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    public class NullValueAttribute : Attribute
    {
        public NullValueAttribute(object value)
        {
            Value = value;
        }

        public object Value { get; }
    }
}