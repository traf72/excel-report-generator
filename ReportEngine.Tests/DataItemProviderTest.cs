using Microsoft.VisualStudio.TestTools.UnitTesting;
using ReportEngine.Exceptions;
using ReportEngine.Implementations.Providers;
using ReportEngine.Tests.Helpers;
using System;

namespace ReportEngine.Tests
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

            var dataItemPropvider = new DataItemValueProvider(dataItem);

            MyAssert.Throws<ArgumentException>(() => dataItemPropvider.GetValue(null));
            MyAssert.Throws<ArgumentException>(() => dataItemPropvider.GetValue(string.Empty));
            MyAssert.Throws<ArgumentException>(() => dataItemPropvider.GetValue(" "));

            Assert.AreEqual(dataItem, dataItemPropvider.GetValue("di"));
            Assert.AreEqual(dataItem.StrProp, dataItemPropvider.GetValue("StrProp"));
            Assert.AreEqual(dataItem.IntProp, dataItemPropvider.GetValue("IntProp"));
            Assert.AreEqual(dataItem.ObjProp, dataItemPropvider.GetValue("ObjProp"));
            Assert.AreEqual(dataItem.ObjProp.StrProp, dataItemPropvider.GetValue("ObjProp.StrProp"));
            Assert.AreEqual(dataItem.ObjProp.ObjProp.GuidProp, dataItemPropvider.GetValue("ObjProp.ObjProp.GuidProp"));

            MyAssert.Throws<MemberNotFoundException>(() => dataItemPropvider.GetValue("DoubleProp"),
                "Cannot find property with name \"DoubleProp\" in class \"TestClass\"");
            MyAssert.Throws<MemberNotFoundException>(() => dataItemPropvider.GetValue("ObjProp.GuidProp"),
                "Cannot find property with name \"GuidProp\" in class \"TestClass2\"");

            dataItem.StrProp = null;
            dataItem.ObjProp = null;

            Assert.IsNull(dataItemPropvider.GetValue("StrProp"));

            MyAssert.Throws<InvalidOperationException>(() => dataItemPropvider.GetValue("ObjProp.StrProp"),
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