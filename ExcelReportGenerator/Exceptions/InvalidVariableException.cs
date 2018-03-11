using System;
using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Exceptions
{
    [LicenceKeyPart(L = true)]
    public class InvalidVariableException : Exception
    {
        public InvalidVariableException(string message) : base(message)
        {
        }

        public InvalidVariableException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}