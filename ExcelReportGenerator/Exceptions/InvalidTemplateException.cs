using System;

namespace ExcelReportGenerator.Exceptions
{
    public class InvalidTemplateException : Exception
    {
        public InvalidTemplateException(string message) : base(message)
        {
        }

        public InvalidTemplateException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}