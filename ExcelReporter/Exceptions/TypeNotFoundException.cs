using System;

namespace ExcelReporter.Exceptions
{
    internal class TypeNotFoundException : Exception
    {
        public TypeNotFoundException(string message) : base(message)
        {
        }

        public TypeNotFoundException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}