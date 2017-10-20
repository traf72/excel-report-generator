using System;

namespace ExcelReporter.Exceptions
{
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