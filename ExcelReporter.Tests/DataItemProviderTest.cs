using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelReporter.Exceptions;
using ExcelReporter.Implementations.Providers;
using ExcelReporter.Tests.Helpers;
using System;

namespace ExcelReporter.Tests
{
    [TestClass]
    public class DataItemProviderTest
    {
        [TestMethod]
        public void TestGetValue()
        {
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
                }
            };

            var dataItemPropvider = new DataItemValueProvider();

            MyAssert.Throws<ArgumentException>(() => dataItemPropvider.GetValue(null, dataItem));
            MyAssert.Throws<ArgumentException>(() => dataItemPropvider.GetValue(string.Empty, dataItem));
            MyAssert.Throws<ArgumentException>(() => dataItemPropvider.GetValue(" ", dataItem));

            Assert.AreEqual(dataItem, dataItemPropvider.GetValue("di", dataItem));
            Assert.AreEqual(dataItem.StrProp, dataItemPropvider.GetValue("StrProp", dataItem));
            Assert.AreEqual(dataItem.IntProp, dataItemPropvider.GetValue("IntProp", dataItem));
            Assert.AreEqual(dataItem.ObjProp, dataItemPropvider.GetValue("ObjProp", dataItem));
            Assert.AreEqual(dataItem.ObjProp.StrProp, dataItemPropvider.GetValue("ObjProp.StrProp", dataItem));
            Assert.AreEqual(dataItem.ObjProp.ObjProp.GuidProp, dataItemPropvider.GetValue("ObjProp.ObjProp.GuidProp", dataItem));

            MyAssert.Throws<MemberNotFoundException>(() => dataItemPropvider.GetValue("DoubleProp", dataItem),
                "Cannot find property with name \"DoubleProp\" in class \"TestClass\"");
            MyAssert.Throws<MemberNotFoundException>(() => dataItemPropvider.GetValue("ObjProp.GuidProp", dataItem),
                "Cannot find property with name \"GuidProp\" in class \"TestClass2\"");

            dataItem.StrProp = null;
            dataItem.ObjProp = null;

            Assert.IsNull(dataItemPropvider.GetValue("StrProp", dataItem));

            MyAssert.Throws<InvalidOperationException>(() => dataItemPropvider.GetValue("ObjProp.StrProp", dataItem),
                "Cannot get property \"StrProp\" because object is null");
        }

        private class TestClass
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
    }
}