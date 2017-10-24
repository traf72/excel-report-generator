using System;

namespace ExcelReporter.Exceptions
{
    internal class IncorrectTemplateException : Exception
    {
        public IncorrectTemplateException(string message) : base(message)
        {
        }

        public IncorrectTemplateException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}