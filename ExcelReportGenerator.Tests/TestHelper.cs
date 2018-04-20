using ClosedXML.Excel;
using System.IO;

namespace ExcelReportGenerator.Tests
{
    public static class TestHelper
    {
        public static XLWorkbook GetExpectedWorkbook(string testClassName, string testMethod)
        {
            return new XLWorkbook(Path.Combine("TestData", testClassName, $"{testMethod}.xlsx"));
        }
    }
}