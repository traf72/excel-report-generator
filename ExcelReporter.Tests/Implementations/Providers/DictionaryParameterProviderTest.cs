using System;
using System.Collections.Generic;
using ExcelReporter.Exceptions;
using ExcelReporter.Implementations.Providers;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Implementations.Providers
{
    [TestClass]
    public class DictionaryParameterProviderTest
    {
        [TestMethod]
        public void TestGetParameterValue()
        {
            ExceptionAssert.Throws<ArgumentNullException>(() => new DictionaryParameterProvider(null));

            var paramsDict = new Dictionary<string, object>
            {
                ["StrParam"] = "value1",
                ["IntParam"] = 5,
                ["BoolParam"] = true,
                ["GuidParam"] = Guid.NewGuid(),
            };

            var provider = new DictionaryParameterProvider(paramsDict);

            Assert.AreEqual(paramsDict["StrParam"], provider.GetParameterValue("StrParam"));
            Assert.AreEqual(paramsDict["IntParam"], provider.GetParameterValue("IntParam"));
            Assert.AreEqual(paramsDict["BoolParam"], provider.GetParameterValue("BoolParam"));
            Assert.AreEqual(paramsDict["GuidParam"], provider.GetParameterValue("GuidParam"));

            ExceptionAssert.Throws<ParameterNotFoundException>(() => provider.GetParameterValue(" StrParam "), "Cannot find paramater with name \" StrParam \"");
            ExceptionAssert.Throws<ParameterNotFoundException>(() => provider.GetParameterValue("strParam"), "Cannot find paramater with name \"strParam\"");
            ExceptionAssert.Throws<ParameterNotFoundException>(() => provider.GetParameterValue("BadParam"), "Cannot find paramater with name \"BadParam\"");
        }
    }
}