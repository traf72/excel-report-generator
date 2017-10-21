using ExcelReporter.Implementations.Providers;
using ExcelReporter.Interfaces.Providers;
using ExcelReporter.Interfaces.TemplateProcessors;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;
using System;
using System.Reflection;
using ExcelReporter.Exceptions;

namespace ExcelReporter.Tests.Implementations.Providers
{
    [TestClass]
    public class MethodCallValueProviderTest
    {
        [TestMethod]
        public void TestParseTemplate()
        {
            var typeProvider = Substitute.For<ITypeProvider>();
            var methodCallValueProvider = new MethodCallValueProvider(typeProvider, null);
            MethodInfo method = methodCallValueProvider.GetType().GetMethod("ParseTemplate", BindingFlags.Instance | BindingFlags.NonPublic);

            object result = method.Invoke(methodCallValueProvider, new[] { "m()" });
            Type resultType = result.GetType();
            PropertyInfo methodNameProp = resultType.GetProperty("MethodName");
            PropertyInfo typeNameProp = resultType.GetProperty("TypeName");
            PropertyInfo methodParamsProp = resultType.GetProperty("MethodParams");

            Assert.AreEqual("m", methodNameProp.GetValue(result));
            Assert.IsNull(typeNameProp.GetValue(result));
            Assert.AreEqual(string.Empty, methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "T:m()" });
            Assert.AreEqual("m", methodNameProp.GetValue(result));
            Assert.AreEqual("T", typeNameProp.GetValue(result));
            Assert.AreEqual(string.Empty, methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { ":T:m()" });
            Assert.AreEqual("m", methodNameProp.GetValue(result));
            Assert.AreEqual(":T", typeNameProp.GetValue(result));
            Assert.AreEqual(string.Empty, methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "ExcelReporter.Tests.Implementations.Providers:T:m()" });
            Assert.AreEqual("m", methodNameProp.GetValue(result));
            Assert.AreEqual("ExcelReporter.Tests.Implementations.Providers:T", typeNameProp.GetValue(result));
            Assert.AreEqual(string.Empty, methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "Method1()" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.IsNull(typeNameProp.GetValue(result));
            Assert.AreEqual(string.Empty, methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "TestClass:Method1()" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.AreEqual("TestClass", typeNameProp.GetValue(result));
            Assert.AreEqual(string.Empty, methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { ":TestClass:Method1()" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.AreEqual(":TestClass", typeNameProp.GetValue(result));
            Assert.AreEqual(string.Empty, methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "ExcelReporter.Tests.Implementations.Providers:TestClass:Method1()" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.AreEqual("ExcelReporter.Tests.Implementations.Providers:TestClass", typeNameProp.GetValue(result));
            Assert.AreEqual(string.Empty, methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "Method1({m:Method2({p:Name})})" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.IsNull(typeNameProp.GetValue(result));
            Assert.AreEqual("{m:Method2({p:Name})}", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "TestClass:Method1({m:Method2({p:Name})})" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.AreEqual("TestClass", typeNameProp.GetValue(result));
            Assert.AreEqual("{m:Method2({p:Name})}", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { ":TestClass:Method1({m:Method2({p:Name})})" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.AreEqual(":TestClass", typeNameProp.GetValue(result));
            Assert.AreEqual("{m:Method2({p:Name})}", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "ExcelReporter.Tests.Implementations.Providers:TestClass:Method1({m:Method2({p:Name})})" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.AreEqual("ExcelReporter.Tests.Implementations.Providers:TestClass", typeNameProp.GetValue(result));
            Assert.AreEqual("{m:Method2({p:Name})}", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "Method1( {p:Name}, {m:Method2({p:Name})}, {di:Field} )" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.IsNull(typeNameProp.GetValue(result));
            Assert.AreEqual("{p:Name}, {m:Method2({p:Name})}, {di:Field}", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "TestClass:Method1( {p:Name}, {m:Method2({p:Name})}, {di:Field} )" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.AreEqual("TestClass", typeNameProp.GetValue(result));
            Assert.AreEqual("{p:Name}, {m:Method2({p:Name})}, {di:Field}", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { ":TestClass:Method1( {p:Name}, {m:Method2({p:Name})}, {di:Field} )" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.AreEqual(":TestClass", typeNameProp.GetValue(result));
            Assert.AreEqual("{p:Name}, {m:Method2({p:Name})}, {di:Field}", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "ExcelReporter.Tests.Implementations.Providers:TestClass:Method1( {p:Name}, {m:Method2({p:Name})}, {di:Field} )" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.AreEqual("ExcelReporter.Tests.Implementations.Providers:TestClass", typeNameProp.GetValue(result));
            Assert.AreEqual("{p:Name}, {m:Method2({p:Name})}, {di:Field}", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "Method1({p:Name}, {m:Method2({p:Name}, { p:value }, {ms:Method3(hi, 5, {p:Desc})})}, {di:Field})" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.IsNull(typeNameProp.GetValue(result));
            Assert.AreEqual("{p:Name}, {m:Method2({p:Name}, { p:value }, {ms:Method3(hi, 5, {p:Desc})})}, {di:Field}", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "TestClass:Method1({p:Name}, {m:Method2({p:Name}, { p:value }, {ms:Method3(hi, 5, {p:Desc})})}, {di:Field})" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.AreEqual("TestClass", typeNameProp.GetValue(result));
            Assert.AreEqual("{p:Name}, {m:Method2({p:Name}, { p:value }, {ms:Method3(hi, 5, {p:Desc})})}, {di:Field}", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { ":TestClass:Method1({p:Name}, {m:Method2({p:Name}, { p:value }, {ms:Method3(hi, 5, {p:Desc})})}, {di:Field})" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.AreEqual(":TestClass", typeNameProp.GetValue(result));
            Assert.AreEqual("{p:Name}, {m:Method2({p:Name}, { p:value }, {ms:Method3(hi, 5, {p:Desc})})}, {di:Field}", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "ExcelReporter.Tests.Implementations.Providers:TestClass:Method1({p:Name}, {m:Method2({p:Name}, { p:value }, {ms:Method3(hi, 5, {p:Desc})})}, {di:Field})" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.AreEqual("ExcelReporter.Tests.Implementations.Providers:TestClass", typeNameProp.GetValue(result));
            Assert.AreEqual("{p:Name}, {m:Method2({p:Name}, { p:value }, {ms:Method3(hi, 5, {p:Desc})})}, {di:Field}", methodParamsProp.GetValue(result));
        }

        [TestMethod]
        public void TestParseParams()
        {
            var typeProvider = Substitute.For<ITypeProvider>();
            var methodCallValueProvider = new MethodCallValueProvider(typeProvider, null);
            MethodInfo method = methodCallValueProvider.GetType().GetMethod("ParseParams", BindingFlags.Instance | BindingFlags.NonPublic);

            var result = (string[])method.Invoke(methodCallValueProvider, new[] { string.Empty });
            Assert.AreEqual(0, result.Length);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { (string)null });
            Assert.AreEqual(0, result.Length);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { " " });
            Assert.AreEqual(0, result.Length);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { "4" });
            Assert.AreEqual(1, result.Length);
            Assert.AreEqual("4", result[0]);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { " {p:Name} " });
            Assert.AreEqual(1, result.Length);
            Assert.AreEqual("{p:Name}", result[0]);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { "{ m:Method() }" });
            Assert.AreEqual(1, result.Length);
            Assert.AreEqual("{ m:Method() }", result[0]);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { "{ms:TestClass:Method()}" });
            Assert.AreEqual(1, result.Length);
            Assert.AreEqual("{ms:TestClass:Method()}", result[0]);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { "{m:Namespace.TestClass:Method()}" });
            Assert.AreEqual(1, result.Length);
            Assert.AreEqual("{m:Namespace.TestClass:Method()}", result[0]);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { "{m::TestClass:Method()}" });
            Assert.AreEqual(1, result.Length);
            Assert.AreEqual("{m::TestClass:Method()}", result[0]);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { "{ms:Method({p:Name})}" });
            Assert.AreEqual(1, result.Length);
            Assert.AreEqual("{ms:Method({p:Name})}", result[0]);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { "{p:Name}, 11, {m:TestClass:Method({p:Name})}, {di:Desc}, hi, {m:Namespace.TestClass:Method({ ms:Method2() })}" });
            Assert.AreEqual(6, result.Length);
            Assert.AreEqual("{p:Name}", result[0]);
            Assert.AreEqual("11", result[1]);
            Assert.AreEqual("{m:TestClass:Method({p:Name})}", result[2]);
            Assert.AreEqual("{di:Desc}", result[3]);
            Assert.AreEqual("hi", result[4]);
            Assert.AreEqual("{m:Namespace.TestClass:Method({ ms:Method2() })}", result[5]);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { "{ m:Method({p:Name}, {di:Desc}) }" });
            Assert.AreEqual(1, result.Length);
            Assert.AreEqual("{ m:Method({p:Name}, {di:Desc}) }", result[0]);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { "{ m:Method({p:Name}, {di:Desc}) }, 11, { m:Method2({m:Namespace.TestClass:Method({ ms:Method2({di:Field}, 777) }, 12, {p:Name})}, 11, {p:Value}) }, ," });
            Assert.AreEqual(5, result.Length);
            Assert.AreEqual("{ m:Method({p:Name}, {di:Desc}) }", result[0]);
            Assert.AreEqual("11", result[1]);
            Assert.AreEqual("{ m:Method2({m:Namespace.TestClass:Method({ ms:Method2({di:Field}, 777) }, 12, {p:Name})}, 11, {p:Value}) }", result[2]);
            Assert.AreEqual(string.Empty, result[3]);
            Assert.AreEqual(string.Empty, result[4]);

            // Экранирование
            result = (string[])method.Invoke(methodCallValueProvider, new[] { "Привет,, медвед!,{{m:Method(10,{p:Name})}}" });
            Assert.AreEqual(2, result.Length);
            Assert.AreEqual("Привет, медвед!", result[0]);
            Assert.AreEqual("{{m:Method(10,{p:Name})}}", result[1]);
        }

        [TestMethod]
        public void TestCallMethod()
        {
            MyAssert.Throws<ArgumentNullException>(() => new MethodCallValueProvider(null, new object()));

            var typeProvider = Substitute.For<ITypeProvider>();
            typeProvider.GetType(Arg.Any<string>()).Returns(typeof(TestClass));

            var templateProcessor = Substitute.For<ITemplateProcessor>();
            var dataItem = new HierarchicalDataItem();
            templateProcessor.Pattern.Returns(@"\{.+?:.+?\}");
            templateProcessor.GetValue("{p:Name}").Returns("TestName");
            templateProcessor.GetValue("{p:Value}", dataItem).Returns(7);

            var methodCallValueProvider = new MethodCallValueProvider(typeProvider, new TestClass());

            MyAssert.Throws<ArgumentException>(() => methodCallValueProvider.CallMethod(null, templateProcessor, new HierarchicalDataItem()));
            MyAssert.Throws<ArgumentException>(() => methodCallValueProvider.CallMethod(string.Empty, templateProcessor, new HierarchicalDataItem()));
            MyAssert.Throws<ArgumentException>(() => methodCallValueProvider.CallMethod(" ", templateProcessor, new HierarchicalDataItem()));

            Assert.AreEqual("Str_1", methodCallValueProvider.CallMethod("Method1()", templateProcessor, null));
            typeProvider.DidNotReceiveWithAnyArgs().GetType(Arg.Any<string>());
            templateProcessor.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());

            typeProvider.ClearReceivedCalls();
            Assert.AreEqual("Str_2", methodCallValueProvider.CallMethod("TestClass:Method1()", templateProcessor, null));
            typeProvider.Received(1).GetType("TestClass");
            templateProcessor.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());

            typeProvider.ClearReceivedCalls();
            Assert.AreEqual("Str_3", methodCallValueProvider.CallMethod(" ExcelReporter.Tests.Implementations.Providers : TestClass : Method1() ", templateProcessor, null));
            typeProvider.Received(1).GetType("ExcelReporter.Tests.Implementations.Providers : TestClass");
            templateProcessor.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());

            typeProvider.ClearReceivedCalls();
            Assert.AreEqual("Str_4", methodCallValueProvider.CallMethod(":TestClass:Method1()", templateProcessor, null));
            typeProvider.Received(1).GetType(":TestClass");
            templateProcessor.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());

            typeProvider.ClearReceivedCalls();
            Assert.AreEqual(25, methodCallValueProvider.CallMethod("Method2({p:Value}, 18)", templateProcessor, dataItem, true));
            typeProvider.DidNotReceiveWithAnyArgs().GetType(Arg.Any<string>());
            templateProcessor.Received(1).GetValue("{p:Value}", dataItem);

            typeProvider.ClearReceivedCalls();
            templateProcessor.ClearReceivedCalls();
            Assert.AreEqual(25, methodCallValueProvider.CallMethod(" : TestClass : Method2({p:Value}, 18) ", templateProcessor, dataItem, true));
            typeProvider.Received(1).GetType(": TestClass");
            templateProcessor.Received(1).GetValue("{p:Value}", dataItem);

            typeProvider.ClearReceivedCalls();
            templateProcessor.ClearReceivedCalls();
            Assert.IsNull(methodCallValueProvider.CallMethod("Method3()", templateProcessor, null));
            typeProvider.DidNotReceiveWithAnyArgs().GetType(Arg.Any<string>());
            templateProcessor.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());

            templateProcessor.GetValue("{ m:TestClass:Method5({p:Desc}, { ms:ExcelReporter.Tests.Implementations.Providers:TestClass:Method6(str, {di.Field})} ) }").Returns(10);
            templateProcessor.GetValue("{m:Method7()}").Returns('c');
            templateProcessor.GetValue("{m::TestClass2:Method1()}").Returns(long.MaxValue);

            object result = methodCallValueProvider.CallMethod(
                "Method4(5, {p:Name}, hi, { m:TestClass:Method5({p:Desc}, { ms:ExcelReporter.Tests.Implementations.Providers:TestClass:Method6(str, {di.Field})} ) }, {m:Method7()}, {m::TestClass2:Method1()})",
                templateProcessor, null);
            Assert.AreEqual($"5_TestName_hi_10_c_{long.MaxValue}", result);
            typeProvider.DidNotReceiveWithAnyArgs().GetType(Arg.Any<string>());
            templateProcessor.Received(1).GetValue("{p:Name}");
            templateProcessor.Received(1).GetValue("{ m:TestClass:Method5({p:Desc}, { ms:ExcelReporter.Tests.Implementations.Providers:TestClass:Method6(str, {di.Field})} ) }");
            templateProcessor.Received(1).GetValue("{m:Method7()}");
            templateProcessor.Received(1).GetValue("{m::TestClass2:Method1()}");

            MyAssert.Throws<MethodNotFoundException>(() => methodCallValueProvider.CallMethod("TestClass:BadMethod()", templateProcessor, null),
                "Could not find public method \"BadMethod\" in type \"TestClass\" and all its parents");

            MyAssert.Throws<MethodNotFoundException>(() => methodCallValueProvider.CallMethod("TestClass:BadMethod()", templateProcessor, null, true),
                "Could not find public static method \"BadMethod\" in type \"TestClass\" and all its parents");

            typeProvider.ClearReceivedCalls();
            templateProcessor.ClearReceivedCalls();
            Assert.AreEqual("Str_Parent", methodCallValueProvider.CallMethod("MethodParent()", templateProcessor, null));
            typeProvider.DidNotReceiveWithAnyArgs().GetType(Arg.Any<string>());
            templateProcessor.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());

            Assert.AreEqual("Str_Static_Parent", methodCallValueProvider.CallMethod("MethodStaticParent()", templateProcessor, null, true));
            typeProvider.DidNotReceiveWithAnyArgs().GetType(Arg.Any<string>());
            templateProcessor.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());
        }

        private class TestClass : TestClassParent
        {
            private int _counter;

            public string Method1()
            {
                _counter++;
                return $"Str_{_counter}";
            }

            public static int Method2(int arg1, string arg2)
            {
                return arg1 + int.Parse(arg2);
            }

            public void Method3()
            {
            }

            public string Method4(string arg1, string arg2, string arg3, int arg4, char arg5, long arg6)
            {
                return $"{arg1}_{arg2}_{arg3}_{arg4}_{arg5}_{arg6}";
            }
        }

        private class TestClassParent
        {
            public string MethodParent()
            {
                return "Str_Parent";
            }

            public static string MethodStaticParent()
            {
                return "Str_Static_Parent";
            }
        }
    }
}
