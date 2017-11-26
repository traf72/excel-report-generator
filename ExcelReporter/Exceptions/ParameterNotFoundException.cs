using System;

namespace ExcelReporter.Exceptions
{
    public class ParameterNotFoundException : Exception
    {
        public ParameterNotFoundException(string message) : base(message)
        {
        }

        public ParameterNotFoundException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}