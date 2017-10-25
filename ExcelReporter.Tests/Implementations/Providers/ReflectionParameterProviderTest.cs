using ExcelReporter.Attributes;
using ExcelReporter.Exceptions;
using ExcelReporter.Implementations.Providers;
using ExcelReporter.Interfaces.Providers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace ExcelReporter.Tests.Implementations.Providers
{
    [TestClass]
    public class ReflectionParameterProviderTest
    {
        [TestMethod]
        public void TestGetParameterValue()
        {
            MyAssert.Throws<ArgumentNullException>(() => new ReflectionParameterProvider(null));

            var complexType = new ComplexParamType();
            Guid guid = Guid.NewGuid();
            var paramContext = new Report
            {
                FieldParameter = "FieldParameter",
                NotParam2 = "NotParam",
                Parameter1 = "Parameter1",
                Parameter2 = 2,
                Parameter3 = complexType,
                ParentParameter = guid,
            };
            IParameterProvider paramProvider = new ReflectionParameterProvider(paramContext);

            Assert.AreEqual("Parameter1", paramProvider.GetParameterValue("Parameter1"));
            Assert.AreEqual("Parameter1", paramProvider.GetParameterValue("Parameter1"));
            Assert.AreEqual("Parameter1", paramProvider.GetParameterValue(" Parameter1 "));
            MyAssert.Throws<ParameterNotFoundException>(() => paramProvider.GetParameterValue("parameter1"), $"Cannot find public instance property or field \"parameter1\" with attribute \"{nameof(Parameter)}\" in type \"Report\" and all its parents");
            Assert.AreSame(complexType, paramProvider.GetParameterValue("Parameter3"));
            Assert.AreEqual("FieldParameter", paramProvider.GetParameterValue("FieldParameter"));
            Assert.AreEqual("FieldParameter", paramProvider.GetParameterValue(" FieldParameter "));
            Assert.AreEqual(guid, paramProvider.GetParameterValue("ParentParameter"));
            MyAssert.Throws<ParameterNotFoundException>(() => paramProvider.GetParameterValue("NotParam1"), $"Cannot find public instance property or field \"NotParam1\" with attribute \"{nameof(Parameter)}\" in type \"Report\" and all its parents");
            MyAssert.Throws<ParameterNotFoundException>(() => paramProvider.GetParameterValue("NotParam2"), $"Cannot find public instance property or field \"NotParam2\" with attribute \"{nameof(Parameter)}\" in type \"Report\" and all its parents");

            MyAssert.Throws<ArgumentException>(() => paramProvider.GetParameterValue(null));
            MyAssert.Throws<ArgumentException>(() => paramProvider.GetParameterValue(string.Empty));
            MyAssert.Throws<ArgumentException>(() => paramProvider.GetParameterValue(" "));
        }

        private class Report : ParentReport
        {
            [Parameter]
            public string Parameter1 { get; set; }

            [Parameter]
            public int Parameter2 { get; set; }

            [Parameter]
            public ComplexParamType Parameter3 { get; set; }

            [Parameter]
            public string FieldParameter;

            [Parameter]
            protected string NotParam1 = "NotParam1";

            public string NotParam2 { get; set; }
        }

        private class ParentReport
        {
            [Parameter]
            public Guid ParentParameter { get; set; }
        }

        private class ComplexParamType
        {
        }
    }
}