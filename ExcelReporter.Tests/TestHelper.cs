using System;
using System.IO;

namespace ExcelReporter.Tests
{
    public static class TestHelper
    {
        public static void InitDataDirectory()
        {
            AppDomain.CurrentDomain.SetData("DataDirectory", new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName);
        }
    }
}