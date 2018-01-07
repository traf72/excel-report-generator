using System;

namespace ExcelReportGenerator.Exceptions
{
    public class IncorrectTemplateException : Exception
    {
        public IncorrectTemplateException(string message) : base(message)
        {
        }

        public IncorrectTemplateException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}