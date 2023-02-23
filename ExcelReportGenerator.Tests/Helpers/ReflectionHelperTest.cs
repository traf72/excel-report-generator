using System.Dynamic;
using System.Reflection;
using ExcelReportGenerator.Attributes;
using ExcelReportGenerator.Exceptions;
using ExcelReportGenerator.Helpers;
using ExcelReportGenerator.Tests.CustomAsserts;

#pragma warning disable 169
#pragma warning disable 108,114

namespace ExcelReportGenerator.Tests.Helpers;

public class ReflectionHelperTest
{
    [Test]
    public void TestGetProperty()
    {
        IReflectionHelper reflectionHelper = new ReflectionHelper();
        Assert.AreEqual("StrProp", reflectionHelper.GetProperty(typeof(TestClass), "StrProp").Name);
        Assert.AreEqual("ObjProp", reflectionHelper.GetProperty(typeof(TestClass), "ObjProp").Name);
        Assert.AreEqual("ParentProp", reflectionHelper.GetProperty(typeof(TestClass), "ParentProp").Name);
        ExceptionAssert.Throws<MemberNotFoundException>(
            () => reflectionHelper.GetProperty(typeof(TestClass), "StaticProp"),
            "Cannot find property \"StaticProp\" in class \"TestClass\" and all its parents. BindingFlags = Instance, Public");
        ExceptionAssert.Throws<MemberNotFoundException>(
            () => reflectionHelper.GetProperty(typeof(TestClass), "PrivateProp"),
            "Cannot find property \"PrivateProp\" in class \"TestClass\" and all its parents. BindingFlags = Instance, Public");
        Assert.AreEqual("StaticProp",
            reflectionHelper.GetProperty(typeof(TestClass), "StaticProp", BindingFlags.Public | BindingFlags.Static)
                .Name);
        Assert.AreEqual("PrivateProp",
            reflectionHelper
                .GetProperty(typeof(TestClass), "PrivateProp", BindingFlags.NonPublic | BindingFlags.Instance).Name);

        var prop = reflectionHelper.GetProperty(typeof(TestClass), "SameNameProp");
        Assert.AreEqual("SameNameProp", prop.Name);
        Assert.AreEqual("ChildSameNameProp", prop.GetValue(new TestClass()));

        prop = reflectionHelper.GetProperty(typeof(TestClass), "SameNameStaticProp",
            BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy);
        Assert.AreEqual("SameNameStaticProp", prop.Name);
        Assert.AreEqual("ChildSameNameStaticProp", prop.GetValue(null));
    }

    [Test]
    public void TestTryGetProperty()
    {
        IReflectionHelper reflectionHelper = new ReflectionHelper();
        Assert.AreEqual("StrProp", reflectionHelper.TryGetProperty(typeof(TestClass), "StrProp").Name);
        Assert.AreEqual("ObjProp", reflectionHelper.TryGetProperty(typeof(TestClass), "ObjProp").Name);
        Assert.AreEqual("ParentProp", reflectionHelper.TryGetProperty(typeof(TestClass), "ParentProp").Name);
        Assert.IsNull(reflectionHelper.TryGetProperty(typeof(TestClass), "StaticProp"));
        Assert.IsNull(reflectionHelper.TryGetProperty(typeof(TestClass), "PrivateProp"));
        Assert.AreEqual("StaticProp",
            reflectionHelper.TryGetProperty(typeof(TestClass), "StaticProp", BindingFlags.Public | BindingFlags.Static)
                .Name);
        Assert.AreEqual("PrivateProp",
            reflectionHelper.TryGetProperty(typeof(TestClass), "PrivateProp",
                BindingFlags.NonPublic | BindingFlags.Instance).Name);

        var prop = reflectionHelper.GetProperty(typeof(TestClass), "SameNameProp");
        Assert.AreEqual("SameNameProp", prop.Name);
        Assert.AreEqual("ChildSameNameProp", prop.GetValue(new TestClass()));

        prop = reflectionHelper.GetProperty(typeof(TestClass), "SameNameStaticProp",
            BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy);
        Assert.AreEqual("SameNameStaticProp", prop.Name);
        Assert.AreEqual("ChildSameNameStaticProp", prop.GetValue(null));
    }

    [Test]
    public void TestGetField()
    {
        IReflectionHelper reflectionHelper = new ReflectionHelper();
        Assert.AreEqual("IntField", reflectionHelper.GetField(typeof(TestClass), "IntField").Name);
        Assert.AreEqual("ParentField", reflectionHelper.GetField(typeof(TestClass), "ParentField").Name);
        ExceptionAssert.Throws<MemberNotFoundException>(
            () => reflectionHelper.GetField(typeof(TestClass), "_privateField"),
            "Cannot find field \"_privateField\" in class \"TestClass\" and all its parents. BindingFlags = Instance, Public");
        ExceptionAssert.Throws<MemberNotFoundException>(
            () => reflectionHelper.GetField(typeof(TestClass), "StaticField"),
            "Cannot find field \"StaticField\" in class \"TestClass\" and all its parents. BindingFlags = Instance, Public");
        Assert.AreEqual("StaticField",
            reflectionHelper.GetField(typeof(TestClass), "StaticField",
                BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy).Name);
        Assert.AreEqual("_privateField",
            reflectionHelper
                .GetField(typeof(TestClass), "_privateField", BindingFlags.NonPublic | BindingFlags.Instance).Name);

        var prop = reflectionHelper.GetField(typeof(TestClass), "SameNameField");
        Assert.AreEqual("SameNameField", prop.Name);
        Assert.AreEqual("ChildSameNameField", prop.GetValue(new TestClass()));

        prop = reflectionHelper.GetField(typeof(TestClass), "SameNameStaticField",
            BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy);
        Assert.AreEqual("SameNameStaticField", prop.Name);
        Assert.AreEqual("ChildSameNameStaticField", prop.GetValue(null));
    }

    [Test]
    public void TestTryGetField()
    {
        IReflectionHelper reflectionHelper = new ReflectionHelper();
        Assert.AreEqual("IntField", reflectionHelper.TryGetField(typeof(TestClass), "IntField").Name);
        Assert.AreEqual("ParentField", reflectionHelper.TryGetField(typeof(TestClass), "ParentField").Name);
        Assert.IsNull(reflectionHelper.TryGetField(typeof(TestClass), "_privateField"));
        Assert.IsNull(reflectionHelper.TryGetField(typeof(TestClass), "StaticField"));
        Assert.AreEqual("StaticField",
            reflectionHelper.TryGetField(typeof(TestClass), "StaticField",
                BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy).Name);
        Assert.AreEqual("_privateField",
            reflectionHelper.TryGetField(typeof(TestClass), "_privateField",
                BindingFlags.NonPublic | BindingFlags.Instance).Name);

        var prop = reflectionHelper.GetField(typeof(TestClass), "SameNameField");
        Assert.AreEqual("SameNameField", prop.Name);
        Assert.AreEqual("ChildSameNameField", prop.GetValue(new TestClass()));

        prop = reflectionHelper.GetField(typeof(TestClass), "SameNameStaticField",
            BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy);
        Assert.AreEqual("SameNameStaticField", prop.Name);
        Assert.AreEqual("ChildSameNameStaticField", prop.GetValue(null));
    }

    [Test]
    public void TestGetValueOfPropertiesChain()
    {
        IReflectionHelper reflectionHelper = new ReflectionHelper();
        var instance = new TestClass();

        Assert.AreSame(instance.StrProp, reflectionHelper.GetValueOfPropertiesChain("StrProp", instance));
        Assert.AreSame(instance.StrProp, reflectionHelper.GetValueOfPropertiesChain(" StrProp ", instance));
        Assert.AreEqual(instance.IntField, reflectionHelper.GetValueOfPropertiesChain("IntField", instance));
        Assert.AreSame(instance.ObjProp, reflectionHelper.GetValueOfPropertiesChain("ObjProp", instance));
        Assert.AreSame(instance.ObjProp.StrProp,
            reflectionHelper.GetValueOfPropertiesChain("ObjProp.StrProp", instance));
        Assert.AreEqual(instance.ObjProp.ObjField.GuidProp,
            reflectionHelper.GetValueOfPropertiesChain("ObjProp.ObjField.GuidProp", instance));
        Assert.AreSame(instance.ParentProp, reflectionHelper.GetValueOfPropertiesChain("ParentProp", instance));
        Assert.AreSame(instance.SameNameProp, reflectionHelper.GetValueOfPropertiesChain("SameNameProp", instance));
        Assert.AreSame(instance.SameNameField, reflectionHelper.GetValueOfPropertiesChain("SameNameField", instance));
        Assert.AreEqual(instance.DynamicObj.GuidProp,
            reflectionHelper.GetValueOfPropertiesChain("DynamicObj.GuidProp", instance));
        Assert.AreEqual(instance.ExpandoObj.StrProp,
            reflectionHelper.GetValueOfPropertiesChain("ExpandoObj.StrProp", instance));
        Assert.AreEqual(instance.ExpandoObj.DecimalProp,
            reflectionHelper.GetValueOfPropertiesChain("ExpandoObj.DecimalProp", instance));
        Assert.AreEqual(instance.ExpandoObj.ComplexProp.GuidProp,
            reflectionHelper.GetValueOfPropertiesChain("ExpandoObj.ComplexProp.GuidProp", instance));
        Assert.AreEqual(instance.ExpandoObj.InnerExpando.IntProp,
            reflectionHelper.GetValueOfPropertiesChain("ExpandoObj.InnerExpando.IntProp", instance));
        Assert.AreEqual(instance.ObjProp.ExpandoField.GuidProp,
            reflectionHelper.GetValueOfPropertiesChain("ObjProp.ExpandoField.GuidProp", instance));

        ExceptionAssert.Throws<MemberNotFoundException>(
            () => reflectionHelper.GetValueOfPropertiesChain("strProp", instance),
            "Cannot find property or field \"strProp\" in class \"TestClass\" and all its parents. BindingFlags = Instance, Public");
        ExceptionAssert.Throws<MemberNotFoundException>(
            () => reflectionHelper.GetValueOfPropertiesChain("StaticProp", instance),
            "Cannot find property or field \"StaticProp\" in class \"TestClass\" and all its parents. BindingFlags = Instance, Public");
        ExceptionAssert.Throws<MemberNotFoundException>(
            () => reflectionHelper.GetValueOfPropertiesChain("_privateField", instance),
            "Cannot find property or field \"_privateField\" in class \"TestClass\" and all its parents. BindingFlags = Instance, Public");
        ExceptionAssert.Throws<MemberNotFoundException>(
            () => reflectionHelper.GetValueOfPropertiesChain("DoubleProp", instance),
            "Cannot find property or field \"DoubleProp\" in class \"TestClass\" and all its parents. BindingFlags = Instance, Public");
        ExceptionAssert.Throws<MemberNotFoundException>(
            () => reflectionHelper.GetValueOfPropertiesChain("ObjProp.GuidProp", instance),
            "Cannot find property or field \"GuidProp\" in class \"TestClass2\" and all its parents. BindingFlags = Instance, Public");
        ExceptionAssert.Throws<MemberNotFoundException>(
            () => reflectionHelper.GetValueOfPropertiesChain("ExpandoObj.BadProp", instance),
            "Cannot find property \"BadProp\" in ExpandoObject");
        ExceptionAssert.Throws<MemberNotFoundException>(
            () => reflectionHelper.GetValueOfPropertiesChain("ExpandoObj.InnerExpando.BadInnerProp", instance),
            "Cannot find property \"BadInnerProp\" in ExpandoObject");

        instance.ObjProp.StrProp = null;
        Assert.AreEqual("DefaultStr", reflectionHelper.GetValueOfPropertiesChain("ObjProp.StrProp", instance));

        instance.StrProp = null;
        instance.ObjProp = null;
        instance.IntField = null;

        Assert.IsNull(reflectionHelper.GetValueOfPropertiesChain("StrProp", instance));
        Assert.AreEqual(777, reflectionHelper.GetValueOfPropertiesChain("IntField", instance));
        Assert.AreEqual("DefaultObjProp", reflectionHelper.GetValueOfPropertiesChain("ObjProp", instance));

        ExceptionAssert.Throws<NullReferenceException>(
            () => reflectionHelper.GetValueOfPropertiesChain("ObjProp.StrProp", instance),
            "Cannot get property or field \"StrProp\" because instance is null");

        ExceptionAssert.Throws<NullReferenceException>(
            () => reflectionHelper.GetValueOfPropertiesChain("IntField", null),
            "Cannot get property or field \"IntField\" because instance is null");

        ExceptionAssert.Throws<ArgumentException>(() => reflectionHelper.GetValueOfPropertiesChain(null, instance));
        ExceptionAssert.Throws<ArgumentException>(() =>
            reflectionHelper.GetValueOfPropertiesChain(string.Empty, instance));
        ExceptionAssert.Throws<ArgumentException>(() => reflectionHelper.GetValueOfPropertiesChain(" ", instance));

        ExceptionAssert.Throws<InvalidOperationException>(
            () => reflectionHelper.GetValueOfPropertiesChain("StrProp", null,
                BindingFlags.Public | BindingFlags.Static),
            "BindingFlags.Static is specified but static properties and fields are not supported");
    }

    [Test]
    public void TestGetNullValueAttributeValue()
    {
        IReflectionHelper reflectionHelper = new ReflectionHelper();
        var instance = new TestClass();

        Assert.AreEqual(777, reflectionHelper.GetNullValueAttributeValue(instance.GetType().GetField("IntField")));
        Assert.AreEqual("DefaultObjProp",
            reflectionHelper.GetNullValueAttributeValue(instance.GetType().GetProperty("ObjProp")));
        Assert.IsNull(reflectionHelper.GetNullValueAttributeValue(instance.GetType().GetProperty("StrProp")));
    }

    private class TestClass : Parent
    {
        public static string SameNameStaticField = "ChildSameNameStaticField";

        private string _privateField;

        [NullValue(777)] public int? IntField = 1;

        public readonly string SameNameField = "ChildSameNameField";

        public TestClass()
        {
            ExpandoObj.StrProp = "Str";
            ExpandoObj.DecimalProp = 56.34m;
            ExpandoObj.ComplexProp = new TestClass3();
            ExpandoObj.InnerExpando = new ExpandoObject();
            ExpandoObj.InnerExpando.IntProp = 100;
        }

        public string StrProp { get; set; } = "StrProp";

        [NullValue("DefaultObjProp")] public TestClass2 ObjProp { get; set; } = new();

        public static string StaticProp { get; set; } = "StaticProp";

        private string PrivateProp { get; set; } = "PrivateProp";

        public string SameNameProp { get; } = "ChildSameNameProp";

        public static string SameNameStaticProp { get; set; } = "ChildSameNameStaticProp";

        public dynamic DynamicObj { get; } = new TestClass3();

        public dynamic ExpandoObj { get; } = new ExpandoObject();
    }

    private class TestClass2
    {
        public readonly TestClass3 ObjField = new();

        public readonly dynamic ExpandoField = new ExpandoObject();

        public TestClass2()
        {
            ExpandoField.GuidProp = Guid.NewGuid();
        }

        [NullValue("DefaultStr")] public string StrProp { get; set; } = "TestClass2:StrProp";
    }

    private class TestClass3
    {
        public Guid GuidProp { get; } = new("5be1d032-6d93-466e-bce0-31dfcefdda22");
    }

    private class Parent
    {
        public static string StaticField = "StaticField";

        public static string SameNameStaticField = "ParentStaticSameNameStaticField";

        public string ParentField = "ParentField";

        public string SameNameField = "ParentSameNameField";
        public string ParentProp { get; } = "ParentProp";

        public string SameNameProp { get; set; } = "ParentSameNameProp";

        public static string SameNameStaticProp { get; set; } = "ParentSameNameStaticProp";
    }
}