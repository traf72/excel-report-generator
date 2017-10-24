using System;

namespace ExcelReporter.Exceptions
{
    internal class ParameterNotFoundException : Exception
    {
        public ParameterNotFoundException(string message) : base(message)
        {
        }

        public ParameterNotFoundException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}