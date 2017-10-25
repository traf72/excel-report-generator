using System;

namespace ExcelReporter.Exceptions
{
    internal class ColumnNotFoundException : Exception
    {
        public ColumnNotFoundException(string message) : base(message)
        {
        }

        public ColumnNotFoundException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}