using ExcelReporter.Exceptions;
using ExcelReporter.Helpers;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Reflection;

#pragma warning disable 169
#pragma warning disable 108,114

namespace ExcelReporter.Tests.Helpers
{
    [TestClass]
    public class ReflectionHelperTest
    {
        [TestMethod]
        public void TestGetProperty()
        {
            IReflectionHelper reflectionHelper = new ReflectionHelper();
            Assert.AreEqual("StrProp", reflectionHelper.GetProperty(typeof(TestClass), "StrProp").Name);
            Assert.AreEqual("ObjProp", reflectionHelper.GetProperty(typeof(TestClass), "ObjProp").Name);
            Assert.AreEqual("ParentProp", reflectionHelper.GetProperty(typeof(TestClass), "ParentProp").Name);
            ExceptionAssert.Throws<MemberNotFoundException>(() => reflectionHelper.GetProperty(typeof(TestClass), "StaticProp"), "Cannot find property \"StaticProp\" in class \"TestClass\" and all its parents. BindingFlags = Instance, Public");
            ExceptionAssert.Throws<MemberNotFoundException>(() => reflectionHelper.GetProperty(typeof(TestClass), "PrivateProp"), "Cannot find property \"PrivateProp\" in class \"TestClass\" and all its parents. BindingFlags = Instance, Public");
            Assert.AreEqual("StaticProp", reflectionHelper.GetProperty(typeof(TestClass), "StaticProp", BindingFlags.Public | BindingFlags.Static).Name);
            Assert.AreEqual("PrivateProp", reflectionHelper.GetProperty(typeof(TestClass), "PrivateProp", BindingFlags.NonPublic | BindingFlags.Instance).Name);

            PropertyInfo prop = reflectionHelper.GetProperty(typeof(TestClass), "SameNameProp");
            Assert.AreEqual("SameNameProp", prop.Name);
            Assert.AreEqual("ChildSameNameProp", prop.GetValue(new TestClass()));

            prop = reflectionHelper.GetProperty(typeof(TestClass), "SameNameStaticProp", BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy);
            Assert.AreEqual("SameNameStaticProp", prop.Name);
            Assert.AreEqual("ChildSameNameStaticProp", prop.GetValue(null));
        }

        [TestMethod]
        public void TestTryGetProperty()
        {
            IReflectionHelper reflectionHelper = new ReflectionHelper();
            Assert.AreEqual("StrProp", reflectionHelper.TryGetProperty(typeof(TestClass), "StrProp").Name);
            Assert.AreEqual("ObjProp", reflectionHelper.TryGetProperty(typeof(TestClass), "ObjProp").Name);
            Assert.AreEqual("ParentProp", reflectionHelper.TryGetProperty(typeof(TestClass), "ParentProp").Name);
            Assert.IsNull(reflectionHelper.TryGetProperty(typeof(TestClass), "StaticProp"));
            Assert.IsNull(reflectionHelper.TryGetProperty(typeof(TestClass), "PrivateProp"));
            Assert.AreEqual("StaticProp", reflectionHelper.TryGetProperty(typeof(TestClass), "StaticProp", BindingFlags.Public | BindingFlags.Static).Name);
            Assert.AreEqual("PrivateProp", reflectionHelper.TryGetProperty(typeof(TestClass), "PrivateProp", BindingFlags.NonPublic | BindingFlags.Instance).Name);

            PropertyInfo prop = reflectionHelper.GetProperty(typeof(TestClass), "SameNameProp");
            Assert.AreEqual("SameNameProp", prop.Name);
            Assert.AreEqual("ChildSameNameProp", prop.GetValue(new TestClass()));

            prop = reflectionHelper.GetProperty(typeof(TestClass), "SameNameStaticProp", BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy);
            Assert.AreEqual("SameNameStaticProp", prop.Name);
            Assert.AreEqual("ChildSameNameStaticProp", prop.GetValue(null));
        }

        [TestMethod]
        public void TestGetField()
        {
            IReflectionHelper reflectionHelper = new ReflectionHelper();
            Assert.AreEqual("IntField", reflectionHelper.GetField(typeof(TestClass), "IntField").Name);
            Assert.AreEqual("ParentField", reflectionHelper.GetField(typeof(TestClass), "ParentField").Name);
            ExceptionAssert.Throws<MemberNotFoundException>(() => reflectionHelper.GetField(typeof(TestClass), "_privateField"), "Cannot find field \"_privateField\" in class \"TestClass\" and all its parents. BindingFlags = Instance, Public");
            ExceptionAssert.Throws<MemberNotFoundException>(() => reflectionHelper.GetField(typeof(TestClass), "StaticField"), "Cannot find field \"StaticField\" in class \"TestClass\" and all its parents. BindingFlags = Instance, Public");
            Assert.AreEqual("StaticField", reflectionHelper.GetField(typeof(TestClass), "StaticField", BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy).Name);
            Assert.AreEqual("_privateField", reflectionHelper.GetField(typeof(TestClass), "_privateField", BindingFlags.NonPublic | BindingFlags.Instance).Name);

            FieldInfo prop = reflectionHelper.GetField(typeof(TestClass), "SameNameField");
            Assert.AreEqual("SameNameField", prop.Name);
            Assert.AreEqual("ChildSameNameField", prop.GetValue(new TestClass()));

            prop = reflectionHelper.GetField(typeof(TestClass), "SameNameStaticField", BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy);
            Assert.AreEqual("SameNameStaticField", prop.Name);
            Assert.AreEqual("ChildSameNameStaticField", prop.GetValue(null));
        }

        [TestMethod]
        public void TestTryGetField()
        {
            IReflectionHelper reflectionHelper = new ReflectionHelper();
            Assert.AreEqual("IntField", reflectionHelper.TryGetField(typeof(TestClass), "IntField").Name);
            Assert.AreEqual("ParentField", reflectionHelper.TryGetField(typeof(TestClass), "ParentField").Name);
            Assert.IsNull(reflectionHelper.TryGetField(typeof(TestClass), "_privateField"));
            Assert.IsNull(reflectionHelper.TryGetField(typeof(TestClass), "StaticField"));
            Assert.AreEqual("StaticField", reflectionHelper.TryGetField(typeof(TestClass), "StaticField", BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy).Name);
            Assert.AreEqual("_privateField", reflectionHelper.TryGetField(typeof(TestClass), "_privateField", BindingFlags.NonPublic | BindingFlags.Instance).Name);

            FieldInfo prop = reflectionHelper.GetField(typeof(TestClass), "SameNameField");
            Assert.AreEqual("SameNameField", prop.Name);
            Assert.AreEqual("ChildSameNameField", prop.GetValue(new TestClass()));

            prop = reflectionHelper.GetField(typeof(TestClass), "SameNameStaticField", BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy);
            Assert.AreEqual("SameNameStaticField", prop.Name);
            Assert.AreEqual("ChildSameNameStaticField", prop.GetValue(null));
        }

        [TestMethod]
        public void TestGetValueOfPropertiesChain()
        {
            IReflectionHelper reflectionHelper = new ReflectionHelper();
            var instance = new TestClass();

            Assert.AreSame(instance.StrProp, reflectionHelper.GetValueOfPropertiesChain("StrProp", instance));
            Assert.AreSame(instance.StrProp, reflectionHelper.GetValueOfPropertiesChain(" StrProp ", instance));
            Assert.AreEqual(instance.IntField, reflectionHelper.GetValueOfPropertiesChain("IntField", instance));
            Assert.AreSame(instance.ObjProp, reflectionHelper.GetValueOfPropertiesChain("ObjProp", instance));
            Assert.AreSame(instance.ObjProp.StrProp, reflectionHelper.GetValueOfPropertiesChain("ObjProp.StrProp", instance));
            Assert.AreEqual(instance.ObjProp.ObjField.GuidProp, reflectionHelper.GetValueOfPropertiesChain("ObjProp.ObjField.GuidProp", instance));
            Assert.AreSame(instance.ParentProp, reflectionHelper.GetValueOfPropertiesChain("ParentProp", instance));
            Assert.AreSame(instance.SameNameProp, reflectionHelper.GetValueOfPropertiesChain("SameNameProp", instance));
            Assert.AreSame(instance.SameNameField, reflectionHelper.GetValueOfPropertiesChain("SameNameField", instance));

            ExceptionAssert.Throws<MemberNotFoundException>(() => reflectionHelper.GetValueOfPropertiesChain("strProp", instance),
                "Cannot find property or field \"strProp\" in class \"TestClass\" and all its parents. BindingFlags = Instance, Public");
            ExceptionAssert.Throws<MemberNotFoundException>(() => reflectionHelper.GetValueOfPropertiesChain("StaticProp", instance),
                "Cannot find property or field \"StaticProp\" in class \"TestClass\" and all its parents. BindingFlags = Instance, Public");
            ExceptionAssert.Throws<MemberNotFoundException>(() => reflectionHelper.GetValueOfPropertiesChain("_privateField", instance),
                "Cannot find property or field \"_privateField\" in class \"TestClass\" and all its parents. BindingFlags = Instance, Public");
            ExceptionAssert.Throws<MemberNotFoundException>(() => reflectionHelper.GetValueOfPropertiesChain("DoubleProp", instance),
                "Cannot find property or field \"DoubleProp\" in class \"TestClass\" and all its parents. BindingFlags = Instance, Public");
            ExceptionAssert.Throws<MemberNotFoundException>(() => reflectionHelper.GetValueOfPropertiesChain("ObjProp.GuidProp", instance),
                "Cannot find property or field \"GuidProp\" in class \"TestClass2\" and all its parents. BindingFlags = Instance, Public");

            instance.StrProp = null;
            instance.ObjProp = null;

            Assert.IsNull(reflectionHelper.GetValueOfPropertiesChain("StrProp", instance));

            ExceptionAssert.Throws<InvalidOperationException>(() => reflectionHelper.GetValueOfPropertiesChain("ObjProp.StrProp", instance),
                "Cannot get property or field \"StrProp\" because instance is null");

            ExceptionAssert.Throws<InvalidOperationException>(() => reflectionHelper.GetValueOfPropertiesChain("IntProp", null),
                "Cannot get property or field \"IntProp\" because instance is null");

            ExceptionAssert.Throws<ArgumentException>(() => reflectionHelper.GetValueOfPropertiesChain(null, instance));
            ExceptionAssert.Throws<ArgumentException>(() => reflectionHelper.GetValueOfPropertiesChain(string.Empty, instance));
            ExceptionAssert.Throws<ArgumentException>(() => reflectionHelper.GetValueOfPropertiesChain(" ", instance));

            ExceptionAssert.Throws<InvalidOperationException>(() => reflectionHelper.GetValueOfPropertiesChain("StrProp", null, BindingFlags.Public | BindingFlags.Static),
                "BindingFlags.Static is specified but static properties and fields are not supported");
        }

        private class TestClass : Parent
        {
            public string StrProp { get; set; } = "StrProp";

            public int IntField = 1;

            public TestClass2 ObjProp { get; set; } = new TestClass2();

            public static string StaticProp { get; set; } = "StaticProp";

            private string PrivateProp { get; set; } = "PrivateProp";

            private string _privateField;

            public string SameNameProp { get; } = "ChildSameNameProp";

            public static string SameNameStaticProp { get; set; } = "ChildSameNameStaticProp";

            public string SameNameField = "ChildSameNameField";

            public static string SameNameStaticField = "ChildSameNameStaticField";
        }

        private class TestClass2
        {
            public string StrProp { get; } = "TestClass2:StrProp";

            public readonly TestClass3 ObjField = new TestClass3();
        }

        private class TestClass3
        {
            public Guid GuidProp { get; } = new Guid("5be1d032-6d93-466e-bce0-31dfcefdda22");
        }

        private class Parent
        {
            public string ParentProp { get; } = "ParentProp";

            public string ParentField = "ParentField";

            public static string StaticField = "StaticField";

            public string SameNameProp { get; set; } = "ParentSameNameProp";

            public static string SameNameStaticProp { get; set; } = "ParentSameNameStaticProp";

            public string SameNameField = "ParentSameNameField";

            public static string SameNameStaticField = "ParentStaticSameNameStaticField";
        }
    }
}