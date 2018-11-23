using System;

namespace ExcelReportGenerator.Exceptions
{
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