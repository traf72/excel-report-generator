using System;
using System.IO;
using ClosedXML.Excel;

namespace ExcelReporter.Tests
{
    public static class TestHelper
    {
        public static void InitDataDirectory()
        {
            AppDomain.CurrentDomain.SetData("DataDirectory", new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName);
        }

        public static XLWorkbook GetExpectedWorkbook(string testClassName, string testMethod)
        {
            return new XLWorkbook(Path.Combine("TestData", testClassName, $"{testMethod}.xlsx"));
        }
    }
}