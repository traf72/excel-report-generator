using System;
using ExcelReporter.Attributes;
using ExcelReporter.Exceptions;
using ExcelReporter.Rendering.Providers.ParameterProviders;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Rendering.Providers.ParameterProviders
{
    [TestClass]
    public class ReflectionParameterProviderTest
    {
        [TestMethod]
        public void TestGetParameterValue()
        {
            ExceptionAssert.Throws<ArgumentNullException>(() => new ReflectionParameterProvider(null));

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
            ExceptionAssert.Throws<ParameterNotFoundException>(() => paramProvider.GetParameterValue("parameter1"), $"Cannot find public instance property or field \"parameter1\" with attribute \"{nameof(ParameterAttribute)}\" in type \"Report\" and all its parents");
            Assert.AreSame(complexType, paramProvider.GetParameterValue("Parameter3"));
            Assert.AreEqual("FieldParameter", paramProvider.GetParameterValue("FieldParameter"));
            Assert.AreEqual("FieldParameter", paramProvider.GetParameterValue(" FieldParameter "));
            Assert.AreEqual(guid, paramProvider.GetParameterValue("ParentParameter"));
            ExceptionAssert.Throws<ParameterNotFoundException>(() => paramProvider.GetParameterValue("NotParam1"), $"Cannot find public instance property or field \"NotParam1\" with attribute \"{nameof(ParameterAttribute)}\" in type \"Report\" and all its parents");
            ExceptionAssert.Throws<ParameterNotFoundException>(() => paramProvider.GetParameterValue("NotParam2"), $"Cannot find public instance property or field \"NotParam2\" with attribute \"{nameof(ParameterAttribute)}\" in type \"Report\" and all its parents");

            ExceptionAssert.Throws<ArgumentException>(() => paramProvider.GetParameterValue(null));
            ExceptionAssert.Throws<ArgumentException>(() => paramProvider.GetParameterValue(string.Empty));
            ExceptionAssert.Throws<ArgumentException>(() => paramProvider.GetParameterValue(" "));
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