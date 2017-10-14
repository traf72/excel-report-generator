using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelReporter.Implementations.Providers;
using ExcelReporter.Implementations.TemplateProcessors;
using System.Collections.Generic;

namespace ExcelReporter.Tests
{
    [TestClass]
    public class TemplateParserTest
    {
        [TestMethod]
        public void TestGetValue()
        {
            var parameterProvider = new DictionaryParameterProvider(new Dictionary<string, object>
            {
                ["Name"] = "TestName",
            });

            var proc = new TemplateProcessor(parameterProvider);
            //object res = proc.Parse("fn:TestFunc(p:Name)");
        }

        public string TestFunc(string param1)
        {
            return $"{param1}: {param1}";
        }
    }
}