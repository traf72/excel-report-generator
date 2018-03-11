using System;
using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Exceptions
{
    [LicenceKeyPart(U = true)]
    public class MethodNotFoundException : Exception
    {
        public MethodNotFoundException(string message) : base(message)
        {
        }

        public MethodNotFoundException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}