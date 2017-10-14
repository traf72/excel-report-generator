using Microsoft.VisualStudio.TestTools.UnitTesting;
using ReportEngine.Implementations.Providers;
using ReportEngine.Implementations.TemplateProcessors;
using System.Collections.Generic;

namespace ReportEngine.Tests
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

            var proc = new TemplateProcessor(parameterProvider, this);
            //object res = proc.Parse("fn:TestFunc(p:Name)");
        }

        public string TestFunc(string param1)
        {
            return $"{param1}: {param1}";
        }
    }
}