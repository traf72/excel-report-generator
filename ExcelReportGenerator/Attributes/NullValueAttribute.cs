using System;

namespace ExcelReportGenerator.Attributes
{
    /// <summary>
    /// An attribute that allows you to replace null-values to more readable
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    [LicenceKeyPart(L = true, R = true)]
    public class NullValueAttribute : Attribute
    {
        public NullValueAttribute(object value)
        {
            Value = value;
        }

        /// <summary>
        /// The value that will be write to Excel if original value is null
        /// </summary>
        public object Value { get; }
    }
}