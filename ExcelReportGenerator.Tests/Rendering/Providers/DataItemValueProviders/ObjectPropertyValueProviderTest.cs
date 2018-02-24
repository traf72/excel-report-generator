using ExcelReportGenerator.Helpers;
using ExcelReportGenerator.Rendering.Providers.DataItemValueProviders;
using ExcelReportGenerator.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;
using System;
using System.Collections.Generic;

namespace ExcelReportGenerator.Tests.Rendering.Providers.DataItemValueProviders
{
    [TestClass]
    public class ObjectPropertyValueProviderTest
    {
        [TestMethod]
        public void TestGetValue()
        {
            var reflectionHelper = Substitute.For<IReflectionHelper>();
            IDataItemValueProvider dataItemValueProvider = new ObjectPropertyValueProvider(reflectionHelper);
            var date = DateTime.Now;

            dataItemValueProvider.GetValue("StrProp", date);
            reflectionHelper.Received(1).GetValueOfPropertiesChain("StrProp", date);

            dataItemValueProvider.GetValue(" StrProp ", date);
            reflectionHelper.Received(2).GetValueOfPropertiesChain("StrProp", date);

            dataItemValueProvider.GetValue("ObjProp.StrProp", date);
            reflectionHelper.Received(1).GetValueOfPropertiesChain("ObjProp.StrProp", date);

            dataItemValueProvider.GetValue("ObjProp.ObjProp.GuidProp", date);
            reflectionHelper.Received(1).GetValueOfPropertiesChain("ObjProp.ObjProp.GuidProp", date);

            ExceptionAssert.Throws<ArgumentException>(() => dataItemValueProvider.GetValue(null, date));
            ExceptionAssert.Throws<ArgumentException>(() => dataItemValueProvider.GetValue(string.Empty, date));
            ExceptionAssert.Throws<ArgumentException>(() => dataItemValueProvider.GetValue(" ", date));
        }

        [TestMethod]
        public void TestGetValueFromKeyValuePair()
        {
            IDataItemValueProvider dataItemValueProvider = new ObjectPropertyValueProvider();
            var dataItem = new KeyValuePair<string, string>("key", "val");
            Assert.AreEqual(dataItem.Key, dataItemValueProvider.GetValue("Key", dataItem));
            Assert.AreEqual(dataItem.Value, dataItemValueProvider.GetValue(" Value ", dataItem));
        }
    }
}