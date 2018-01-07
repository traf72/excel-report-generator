using System;
using System.Reflection;
using ExcelReporter.Attributes;
using ExcelReporter.Exceptions;
using ExcelReporter.Helpers;
using ExcelReporter.Rendering.Providers;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;

namespace ExcelReporter.Tests.Rendering.Providers
{
    [TestClass]
    public class DefaultPropertyValueProviderTest
    {
        [TestMethod]
        public void TestParseTemplate()
        {
            var typeProvider = Substitute.For<ITypeProvider>();
            var instanceProvider = Substitute.For<IInstanceProvider>();
            var propertyValueProvider = new DefaultPropertyValueProvider(typeProvider, instanceProvider);
            MethodInfo method = propertyValueProvider.GetType().GetMethod("ParseTemplate", BindingFlags.Instance | BindingFlags.NonPublic);

            object result = method.Invoke(propertyValueProvider, new[] {"Prop"});
            Type resultType = result.GetType();
            PropertyInfo propNameProp = resultType.GetProperty("MemberName");
            PropertyInfo typeNameProp = resultType.GetProperty("TypeName");

            Assert.AreEqual("Prop", propNameProp.GetValue(result));
            Assert.IsNull(typeNameProp.GetValue(result));

            result = method.Invoke(propertyValueProvider, new[] { "Prop.ParentProp.ParentProp" });
            Assert.AreEqual("Prop.ParentProp.ParentProp", propNameProp.GetValue(result));
            Assert.IsNull(typeNameProp.GetValue(result));

            result = method.Invoke(propertyValueProvider, new[] { "T:Prop" });
            Assert.AreEqual("Prop", propNameProp.GetValue(result));
            Assert.AreEqual("T", typeNameProp.GetValue(result));

            result = method.Invoke(propertyValueProvider, new[] { ":T:Prop" });
            Assert.AreEqual("Prop", propNameProp.GetValue(result));
            Assert.AreEqual(":T", typeNameProp.GetValue(result));

            result = method.Invoke(propertyValueProvider, new[] { "T:Prop.ParentProp.ParentProp" });
            Assert.AreEqual("Prop.ParentProp.ParentProp", propNameProp.GetValue(result));
            Assert.AreEqual("T", typeNameProp.GetValue(result));

            result = method.Invoke(propertyValueProvider, new[] { " ExcelReporter.Tests.Implementations.Providers:T : Prop " });
            Assert.AreEqual("Prop", propNameProp.GetValue(result));
            Assert.AreEqual("ExcelReporter.Tests.Implementations.Providers:T", typeNameProp.GetValue(result));

            result = method.Invoke(propertyValueProvider, new[] { "ExcelReporter.Tests.Implementations.Providers:T:Prop.ParentProp.ParentProp" });
            Assert.AreEqual("Prop.ParentProp.ParentProp", propNameProp.GetValue(result));
            Assert.AreEqual("ExcelReporter.Tests.Implementations.Providers:T", typeNameProp.GetValue(result));
        }

        [TestMethod]
        public void TestGetValue()
        {
            ExceptionAssert.Throws<ArgumentNullException>(() => new DefaultPropertyValueProvider(null, Substitute.For<IInstanceProvider>()));
            ExceptionAssert.Throws<ArgumentNullException>(() => new DefaultPropertyValueProvider(Substitute.For<ITypeProvider>(), null));

            var typeProvider = new DefaultTypeProvider(new[] { Assembly.GetExecutingAssembly(), Assembly.GetAssembly(typeof(DateTime)), }, typeof(PropValProviderTestClass));
            var testInstance = new PropValProviderTestClass();
            var instanceProvider = new DefaultInstanceProvider(testInstance);
            var reflectionHelper = new ReflectionHelper();

            IPropertyValueProvider propertyValueProvider = new DefaultPropertyValueProvider(typeProvider, instanceProvider, reflectionHelper);
            Assert.AreEqual(testInstance.StrProp, propertyValueProvider.GetValue("StrProp"));
            Assert.AreEqual(testInstance.StrProp, propertyValueProvider.GetValue("PropValProviderTestClass: StrProp "));
            Assert.AreEqual(testInstance.IntField, propertyValueProvider.GetValue("IntField"));
            Assert.AreEqual(PropValProviderTestClass.StaticStrProp, propertyValueProvider.GetValue("StaticStrProp"));
            Assert.AreEqual(PropValProviderTestClass.StaticIntField, propertyValueProvider.GetValue("StaticIntField"));
            Assert.AreEqual(testInstance.ParentProp, propertyValueProvider.GetValue("ParentProp"));
            Assert.AreEqual(testInstance.ParentField, propertyValueProvider.GetValue("ExcelReporter.Tests.Rendering.Providers:PropValProviderTestClass:ParentField"));
            Assert.AreEqual(PropValProviderTestClass.StaticParentProp, propertyValueProvider.GetValue("StaticParentProp"));
            Assert.AreEqual(Parent.StaticParentField, propertyValueProvider.GetValue("StaticParentField"));

            Assert.AreEqual("TestClass2:StrProp", propertyValueProvider.GetValue("ObjProp.StrProp"));
            Assert.AreEqual("TestClass2:StrProp", propertyValueProvider.GetValue("ExcelReporter.Tests.Rendering.Providers:PropValProviderTestClass:ObjField.StrProp"));
            Assert.AreEqual("TestClass2:StrField", propertyValueProvider.GetValue("PropValProviderTestClass:ObjProp.StrField"));
            Assert.AreEqual("TestClass2:StrField", propertyValueProvider.GetValue("ObjField.StrField"));

            Assert.AreEqual("TestClass2:StrProp", propertyValueProvider.GetValue("PropValProviderTestClass:StaticObjProp.StrProp"));
            Assert.AreEqual("TestClass2:StrProp", propertyValueProvider.GetValue("StaticObjField.StrProp"));

            Assert.AreEqual(Guid.Parse("5be1d032-6d93-466e-bce0-31dfcefdda22"), propertyValueProvider.GetValue("ObjProp.ObjField.GuidProp"));
            Assert.AreEqual(Guid.Parse("5be1d032-6d93-466e-bce0-31dfcefdda22"), propertyValueProvider.GetValue("ObjField.ObjField.GuidProp"));

            Assert.AreEqual(Guid.Parse("5be1d032-6d93-466e-bce0-31dfcefdda22"), propertyValueProvider.GetValue("StaticObjProp.ObjField.GuidProp"));
            Assert.AreEqual(Guid.Parse("5be1d032-6d93-466e-bce0-31dfcefdda22"), propertyValueProvider.GetValue("ExcelReporter.Tests.Rendering.Providers:PropValProviderTestClass:StaticObjField.ObjField.GuidProp"));

            testInstance.StrProp = null;
            Assert.AreEqual("DefaultStrProp", propertyValueProvider.GetValue("StrProp"));

            testInstance.ObjProp.StrField = null;
            Assert.AreEqual("DefaultStrField", propertyValueProvider.GetValue("ObjProp.StrField"));

            testInstance.ObjProp.ObjField.GuidProp = null;
            Assert.AreEqual(0, propertyValueProvider.GetValue("ObjProp.ObjField.GuidProp"));

            testInstance.ObjProp = null;
            Assert.IsNull(propertyValueProvider.GetValue("ObjProp"));

            ExceptionAssert.Throws<ArgumentException>(() => propertyValueProvider.GetValue(null));
            ExceptionAssert.Throws<ArgumentException>(() => propertyValueProvider.GetValue(string.Empty));
            ExceptionAssert.Throws<ArgumentException>(() => propertyValueProvider.GetValue(" "));

            ExceptionAssert.Throws<IncorrectTemplateException>(() => propertyValueProvider.GetValue("PropValProviderTestClass:"));
            ExceptionAssert.Throws<IncorrectTemplateException>(() => propertyValueProvider.GetValue("ExcelReporter.Tests.Rendering.Providers:PropValProviderTestClass: "));

            ExceptionAssert.Throws<MemberNotFoundException>(() => propertyValueProvider.GetValue("BadField"),
                "Cannot find property or field \"BadField\" in class \"PropValProviderTestClass\" and all its parents. BindingFlags = Instance, Static, Public, FlattenHierarchy");

            ExceptionAssert.Throws<InvalidOperationException>(() => propertyValueProvider.GetValue("PrivateGetterProp"),
                "Property \"PrivateGetterProp\" of type \"PropValProviderTestClass\" has no public getter");

            // Static system type props
            Assert.AreEqual(DateTime.Now.Date, ((DateTime)propertyValueProvider.GetValue("DateTime:Now.Date")));
            Assert.AreEqual(Guid.Empty, (Guid)propertyValueProvider.GetValue("Guid:Empty"));

            // Non static system type prop
            Assert.AreEqual(0, propertyValueProvider.GetValue("Version:Minor"));
        }

        private class PropValProviderTestClass : Parent
        {
            [NullValue("DefaultStrProp")]
            public string StrProp { get; set; } = "StrProp";

            public int IntField = 1;

            public static string StaticStrProp { get; set; } = "StaticStrProp";

            public static int StaticIntField = 1000;

            public TestClass2 ObjProp { get; set; } = new TestClass2();

            public TestClass2 ObjField = new TestClass2();

            public static TestClass2 StaticObjProp { get; set; } = new TestClass2();

            public static TestClass2 StaticObjField = new TestClass2();

            public string PrivateGetterProp { private get; set; } = "PrivateGetterProp";
        }

        private class TestClass2
        {
            public string StrProp { get; set; } = "TestClass2:StrProp";

            [NullValue("DefaultStrField")]
            public string StrField = "TestClass2:StrField";

            public TestClass3 ObjField = new TestClass3();
        }

        private class TestClass3
        {
            [NullValue(0)]
            public Guid? GuidProp { get; set; } = new Guid("5be1d032-6d93-466e-bce0-31dfcefdda22");
        }

        private class Parent
        {
            public string ParentProp { get; set; } = "ParentProp";

            public static string StaticParentProp { get; set; } = "StaticParentProp";

            public string ParentField = "ParentField";

            public static string StaticParentField = "StaticParentField";
        }
    }
}