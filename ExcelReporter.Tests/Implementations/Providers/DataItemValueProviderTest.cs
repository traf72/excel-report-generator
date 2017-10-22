using ExcelReporter.Exceptions;
using ExcelReporter.Implementations.Providers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace ExcelReporter.Tests.Implementations.Providers
{
    [TestClass]
    public class DataItemValueProviderTest
    {
        [TestMethod]
        public void TestGetValue()
        {
            var dataItemValueProvider = new DataItemValueProvider();
            var dataItem1 = new TestClass
            {
                HierarchyLevel = 1,
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

            var dataItem2 = new TestClass2
            {
                HierarchyLevel = 2,
                StrProp = "Str",
                ObjProp = new TestClass3
                {
                    GuidProp = Guid.NewGuid()
                }
            };

            var dataItem3 = new TestClass3
            {
                HierarchyLevel = 3,
                GuidProp = Guid.NewGuid()
            };

            var hierarchicalDataItem = new HierarchicalDataItem
            {
                Value = dataItem1,
                Parent = new HierarchicalDataItem
                {
                    Value = dataItem2,
                    Parent = new HierarchicalDataItem
                    {
                        Value = dataItem3,
                    }
                }
            };

            MyAssert.Throws<ArgumentException>(() => dataItemValueProvider.GetValue(null, hierarchicalDataItem));
            MyAssert.Throws<ArgumentException>(() => dataItemValueProvider.GetValue(string.Empty, hierarchicalDataItem));
            MyAssert.Throws<ArgumentException>(() => dataItemValueProvider.GetValue(" ", hierarchicalDataItem));

            MyAssert.Throws<ArgumentNullException>(() => dataItemValueProvider.GetValue("Template", null));
            MyAssert.Throws<ArgumentNullException>(() => dataItemValueProvider.GetValue("Template", null));
            MyAssert.Throws<ArgumentNullException>(() => dataItemValueProvider.GetValue("Template", null));

            Assert.AreEqual(dataItem1, dataItemValueProvider.GetValue("di", hierarchicalDataItem));
            Assert.AreEqual(dataItem1.StrProp, dataItemValueProvider.GetValue("StrProp", hierarchicalDataItem));
            Assert.AreEqual(dataItem1.IntProp, dataItemValueProvider.GetValue("IntProp", hierarchicalDataItem));
            Assert.AreEqual(dataItem1.ObjProp, dataItemValueProvider.GetValue("ObjProp", hierarchicalDataItem));
            Assert.AreEqual(dataItem1.HierarchyLevel, dataItemValueProvider.GetValue("HierarchyLevel", hierarchicalDataItem));
            Assert.AreEqual(dataItem1.ObjProp.StrProp, dataItemValueProvider.GetValue("ObjProp.StrProp", hierarchicalDataItem));
            Assert.AreEqual(dataItem1.ObjProp.HierarchyLevel, dataItemValueProvider.GetValue("ObjProp.HierarchyLevel", hierarchicalDataItem));
            Assert.AreEqual(dataItem1.ObjProp.ObjProp.GuidProp, dataItemValueProvider.GetValue("ObjProp.ObjProp.GuidProp", hierarchicalDataItem));
            Assert.AreEqual(dataItem1.ObjProp.ObjProp.HierarchyLevel, dataItemValueProvider.GetValue("ObjProp.ObjProp.HierarchyLevel", hierarchicalDataItem));
            Assert.AreEqual(dataItem1.ParentProp, dataItemValueProvider.GetValue("ParentProp", hierarchicalDataItem));

            MyAssert.Throws<MemberNotFoundException>(() => dataItemValueProvider.GetValue("DoubleProp", hierarchicalDataItem),
                "Cannot find public instance property \"DoubleProp\" in class \"TestClass\" and all its parents");
            MyAssert.Throws<MemberNotFoundException>(() => dataItemValueProvider.GetValue("ObjProp.GuidProp", hierarchicalDataItem),
                "Cannot find public instance property \"GuidProp\" in class \"TestClass2\" and all its parents");

            dataItem1.StrProp = null;
            dataItem1.ObjProp = null;

            Assert.IsNull(dataItemValueProvider.GetValue("StrProp", hierarchicalDataItem));

            MyAssert.Throws<InvalidOperationException>(() => dataItemValueProvider.GetValue("ObjProp.StrProp", hierarchicalDataItem),
                "Cannot get property \"StrProp\" because object is null");

            Assert.AreEqual(dataItem2, dataItemValueProvider.GetValue("parent:di", hierarchicalDataItem));
            Assert.AreEqual(dataItem2.StrProp, dataItemValueProvider.GetValue("parent:StrProp", hierarchicalDataItem));
            Assert.AreEqual(dataItem2.HierarchyLevel, dataItemValueProvider.GetValue("parent:HierarchyLevel", hierarchicalDataItem));
            Assert.AreEqual(dataItem2.ObjProp.GuidProp, dataItemValueProvider.GetValue("parent:ObjProp.GuidProp", hierarchicalDataItem));
            Assert.AreEqual(dataItem2.ObjProp.HierarchyLevel, dataItemValueProvider.GetValue("parent:ObjProp.HierarchyLevel", hierarchicalDataItem));

            Assert.AreEqual(dataItem3, dataItemValueProvider.GetValue("parent:parent:di", hierarchicalDataItem));
            Assert.AreEqual(dataItem3.GuidProp, dataItemValueProvider.GetValue("parent:Parent:GuidProp", hierarchicalDataItem));
            Assert.AreEqual(dataItem3.HierarchyLevel, dataItemValueProvider.GetValue("parent : PARENT : HierarchyLevel", hierarchicalDataItem));

            MyAssert.Throws<InvalidOperationException>(() => dataItemValueProvider.GetValue("parent:parent:parent:HierarchyLevel", hierarchicalDataItem),
                "Data item is null for template \"parent:parent:parent:HierarchyLevel\"");
            MyAssert.Throws<IncorrectTemplateException>(() => dataItemValueProvider.GetValue("parent:bad:HierarchyLevel", hierarchicalDataItem),
                "Template \"parent:bad:HierarchyLevel\" is incorrect");
        }

        private class TestClass : Parent
        {
            public string StrProp { get; set; }

            public int IntProp { get; set; }

            public TestClass2 ObjProp { get; set; }
        }

        private class TestClass2
        {
            public int HierarchyLevel { get; set; }

            public string StrProp { get; set; }

            public TestClass3 ObjProp { get; set; }
        }

        private class TestClass3
        {
            public int HierarchyLevel { get; set; }

            public Guid GuidProp { get; set; }
        }

        private class Parent
        {
            public int HierarchyLevel { get; set; }

            public string ParentProp { get; set; }
        }
    }
}