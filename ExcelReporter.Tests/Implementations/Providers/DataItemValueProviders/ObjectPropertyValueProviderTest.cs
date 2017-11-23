using ExcelReporter.Exceptions;
using ExcelReporter.Implementations.Providers.DataItemValueProviders;
using ExcelReporter.Interfaces.Providers.DataItemValueProviders;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace ExcelReporter.Tests.Implementations.Providers.DataItemValueProviders
{
    [TestClass]
    public class ObjectPropertyValueProviderTest
    {
        [TestMethod]
        public void TestGetValue()
        {
            IDataItemValueProvider dataItemValueProvider = new ObjectPropertyValueProvider();
            var dataItem = new TestClass
            {
                IntProp = 5,
                StrProp = "Str",
                ObjProp = new TestClass2
                {
                    StrProp = "Str2",
                    ObjProp = new TestClass3
                    {
                        GuidProp = Guid.NewGuid()
                    }
                },
                ParentProp = "Parent",
            };

            Assert.AreEqual(dataItem, dataItemValueProvider.GetValue("di", dataItem));
            Assert.IsNull(dataItemValueProvider.GetValue(" di ", null));
            Assert.AreEqual(dataItem.StrProp, dataItemValueProvider.GetValue("StrProp", dataItem));
            Assert.AreEqual(dataItem.StrProp, dataItemValueProvider.GetValue(" StrProp ", dataItem));
            Assert.AreEqual(dataItem.IntProp, dataItemValueProvider.GetValue("IntProp", dataItem));
            Assert.AreEqual(dataItem.ObjProp, dataItemValueProvider.GetValue("ObjProp", dataItem));
            Assert.AreEqual(dataItem.ObjProp.StrProp, dataItemValueProvider.GetValue("ObjProp.StrProp", dataItem));
            Assert.AreEqual(dataItem.ObjProp.ObjProp.GuidProp, dataItemValueProvider.GetValue("ObjProp.ObjProp.GuidProp", dataItem));
            Assert.AreEqual(dataItem.ParentProp, dataItemValueProvider.GetValue("ParentProp", dataItem));

            ExceptionAssert.Throws<MemberNotFoundException>(() => dataItemValueProvider.GetValue("strProp", dataItem),
                "Cannot find public instance property \"strProp\" in class \"TestClass\" and all its parents");
            ExceptionAssert.Throws<MemberNotFoundException>(() => dataItemValueProvider.GetValue("DoubleProp", dataItem),
                "Cannot find public instance property \"DoubleProp\" in class \"TestClass\" and all its parents");
            ExceptionAssert.Throws<MemberNotFoundException>(() => dataItemValueProvider.GetValue("ObjProp.GuidProp", dataItem),
                "Cannot find public instance property \"GuidProp\" in class \"TestClass2\" and all its parents");

            dataItem.StrProp = null;
            dataItem.ObjProp = null;

            Assert.IsNull(dataItemValueProvider.GetValue("StrProp", dataItem));

            ExceptionAssert.Throws<InvalidOperationException>(() => dataItemValueProvider.GetValue("ObjProp.StrProp", dataItem),
                "Cannot get property \"StrProp\" because object is null");

            ExceptionAssert.Throws<InvalidOperationException>(() => dataItemValueProvider.GetValue("IntProp", null),
                "Cannot get property \"IntProp\" because object is null");

            ExceptionAssert.Throws<ArgumentException>(() => dataItemValueProvider.GetValue(null, dataItem));
            ExceptionAssert.Throws<ArgumentException>(() => dataItemValueProvider.GetValue(string.Empty, dataItem));
            ExceptionAssert.Throws<ArgumentException>(() => dataItemValueProvider.GetValue(" ", dataItem));
        }

        [TestMethod]
        public void TestGetValueFromKeyValuePair()
        {
            IDataItemValueProvider dataItemValueProvider = new ObjectPropertyValueProvider();
            var dataItem = new KeyValuePair<string, string>("key", "val");
            Assert.AreEqual(dataItem, dataItemValueProvider.GetValue("di", dataItem));
            Assert.AreEqual(dataItem.Key, dataItemValueProvider.GetValue("Key", dataItem));
            Assert.AreEqual(dataItem.Value, dataItemValueProvider.GetValue(" Value ", dataItem));
        }

        private class TestClass : Parent
        {
            public string StrProp { get; set; }

            public int IntProp { get; set; }

            public TestClass2 ObjProp { get; set; }
        }

        private class TestClass2
        {
            public string StrProp { get; set; }

            public TestClass3 ObjProp { get; set; }
        }

        private class TestClass3
        {
            public Guid GuidProp { get; set; }
        }

        private class Parent
        {
            public string ParentProp { get; set; }
        }
    }
}