using System;

namespace ExcelReportGenerator.Exceptions
{
    public class MemberNotFoundException : Exception
    {
        public MemberNotFoundException(string message) : base(message)
        {
        }

        public MemberNotFoundException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}