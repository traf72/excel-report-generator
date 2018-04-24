using System;
using System.Collections.Generic;
using ExcelReportGenerator.Rendering.Providers.DataItemValueProviders;
using ExcelReportGenerator.Tests.CustomAsserts;
using NUnit.Framework;

namespace ExcelReportGenerator.Tests.Rendering.Providers.DataItemValueProviders
{
    
    public class DictionaryValueProviderTest
    {
        [Test]
        public void TestGetValue()
        {
            var dict = new Dictionary<string, object>
            {
                ["StrParam"] = "value1",
                ["IntParam"] = 5,
                ["BoolParam"] = true,
                ["GuidParam"] = Guid.NewGuid(),
            };

            var provider = new DictionaryValueProvider<object>();

            Assert.AreEqual(dict["StrParam"], provider.GetValue("StrParam", dict));
            Assert.AreEqual(dict["IntParam"], provider.GetValue("IntParam", dict));
            Assert.AreEqual(dict["BoolParam"], provider.GetValue("BoolParam", dict));
            Assert.AreEqual(dict["GuidParam"], provider.GetValue("GuidParam", dict));

            ExceptionAssert.Throws<ArgumentException>(() => provider.GetValue(null, dict));
            ExceptionAssert.Throws<KeyNotFoundException>(() => provider.GetValue(" StrParam ", dict), "Key \" StrParam \" was not found in dictionary");
            ExceptionAssert.Throws<KeyNotFoundException>(() => provider.GetValue("strParam", dict), "Key \"strParam\" was not found in dictionary");
            ExceptionAssert.Throws<KeyNotFoundException>(() => provider.GetValue("BadParam", dict), "Key \"BadParam\" was not found in dictionary");
        }

        [Test]
        public void TestGetValueIfDictionaryHasDecimalValues()
        {
            var dict = new Dictionary<string, decimal>
            {
                ["Key1"] = 27.67m,
                ["Key2"] = 64m,
            };

            var provider = new DictionaryValueProvider<decimal>();

            Assert.AreEqual(dict["Key1"], provider.GetValue("Key1", dict));
            Assert.AreEqual(dict["Key2"], provider.GetValue("Key2", dict));
        }
    }
}