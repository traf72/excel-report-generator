using ExcelReporter.Implementations.Providers;
using ExcelReporter.Interfaces.Providers;
using ExcelReporter.Interfaces.TemplateProcessors;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;
using System;
using System.Reflection;

namespace ExcelReporter.Tests.Implementations.Providers
{
    [TestClass]
    public class MethodCallValueProviderTest
    {
        [TestMethod]
        public void TestParseTemplate()
        {
            var typeProvider = Substitute.For<ITypeProvider>();
            var methodCallValueProvider = new MethodCallValueProvider(typeProvider);
            MethodInfo method = methodCallValueProvider.GetType().GetMethod("ParseTemplate", BindingFlags.Instance | BindingFlags.NonPublic);

            var result = (MethodCallTemplateParts)method.Invoke(methodCallValueProvider, new[] { "m()" });
            Assert.AreEqual("m", result.MethodName);
            Assert.IsNull(result.TypeName);
            Assert.AreEqual(string.Empty, result.MethodParams);

            result = (MethodCallTemplateParts)method.Invoke(methodCallValueProvider, new[] { "T:m()" });
            Assert.AreEqual("m", result.MethodName);
            Assert.AreEqual("T", result.TypeName);
            Assert.AreEqual(string.Empty, result.MethodParams);

            result = (MethodCallTemplateParts)method.Invoke(methodCallValueProvider, new[] { ":T:m()" });
            Assert.AreEqual("m", result.MethodName);
            Assert.AreEqual(":T", result.TypeName);
            Assert.AreEqual(string.Empty, result.MethodParams);

            result = (MethodCallTemplateParts)method.Invoke(methodCallValueProvider, new[] { "ExcelReporter.Tests.Implementations.Providers:T:m()" });
            Assert.AreEqual("m", result.MethodName);
            Assert.AreEqual("ExcelReporter.Tests.Implementations.Providers:T", result.TypeName);
            Assert.AreEqual(string.Empty, result.MethodParams);

            result = (MethodCallTemplateParts)method.Invoke(methodCallValueProvider, new[] { "Method1()" });
            Assert.AreEqual("Method1", result.MethodName);
            Assert.IsNull(result.TypeName);
            Assert.AreEqual(string.Empty, result.MethodParams);

            result = (MethodCallTemplateParts)method.Invoke(methodCallValueProvider, new[] { "TestClass:Method1()" });
            Assert.AreEqual("Method1", result.MethodName);
            Assert.AreEqual("TestClass", result.TypeName);
            Assert.AreEqual(string.Empty, result.MethodParams);

            result = (MethodCallTemplateParts)method.Invoke(methodCallValueProvider, new[] { ":TestClass:Method1()" });
            Assert.AreEqual("Method1", result.MethodName);
            Assert.AreEqual(":TestClass", result.TypeName);
            Assert.AreEqual(string.Empty, result.MethodParams);

            result = (MethodCallTemplateParts)method.Invoke(methodCallValueProvider, new[] { "ExcelReporter.Tests.Implementations.Providers:TestClass:Method1()" });
            Assert.AreEqual("Method1", result.MethodName);
            Assert.AreEqual("ExcelReporter.Tests.Implementations.Providers:TestClass", result.TypeName);
            Assert.AreEqual(string.Empty, result.MethodParams);

            result = (MethodCallTemplateParts)method.Invoke(methodCallValueProvider, new[] { "Method1({m:Method2({p:Name})})" });
            Assert.AreEqual("Method1", result.MethodName);
            Assert.IsNull(result.TypeName);
            Assert.AreEqual("{m:Method2({p:Name})}", result.MethodParams);

            result = (MethodCallTemplateParts)method.Invoke(methodCallValueProvider, new[] { "TestClass:Method1({m:Method2({p:Name})})" });
            Assert.AreEqual("Method1", result.MethodName);
            Assert.AreEqual("TestClass", result.TypeName);
            Assert.AreEqual("{m:Method2({p:Name})}", result.MethodParams);

            result = (MethodCallTemplateParts)method.Invoke(methodCallValueProvider, new[] { ":TestClass:Method1({m:Method2({p:Name})})" });
            Assert.AreEqual("Method1", result.MethodName);
            Assert.AreEqual(":TestClass", result.TypeName);
            Assert.AreEqual("{m:Method2({p:Name})}", result.MethodParams);

            result = (MethodCallTemplateParts)method.Invoke(methodCallValueProvider, new[] { "ExcelReporter.Tests.Implementations.Providers:TestClass:Method1({m:Method2({p:Name})})" });
            Assert.AreEqual("Method1", result.MethodName);
            Assert.AreEqual("ExcelReporter.Tests.Implementations.Providers:TestClass", result.TypeName);
            Assert.AreEqual("{m:Method2({p:Name})}", result.MethodParams);

            result = (MethodCallTemplateParts)method.Invoke(methodCallValueProvider, new[] { "Method1( {p:Name}, {m:Method2({p:Name})}, {di:Field} )" });
            Assert.AreEqual("Method1", result.MethodName);
            Assert.IsNull(result.TypeName);
            Assert.AreEqual(" {p:Name}, {m:Method2({p:Name})}, {di:Field} ", result.MethodParams);

            result = (MethodCallTemplateParts)method.Invoke(methodCallValueProvider, new[] { "TestClass:Method1( {p:Name}, {m:Method2({p:Name})}, {di:Field} )" });
            Assert.AreEqual("Method1", result.MethodName);
            Assert.AreEqual("TestClass", result.TypeName);
            Assert.AreEqual(" {p:Name}, {m:Method2({p:Name})}, {di:Field} ", result.MethodParams);

            result = (MethodCallTemplateParts)method.Invoke(methodCallValueProvider, new[] { ":TestClass:Method1( {p:Name}, {m:Method2({p:Name})}, {di:Field} )" });
            Assert.AreEqual("Method1", result.MethodName);
            Assert.AreEqual(":TestClass", result.TypeName);
            Assert.AreEqual(" {p:Name}, {m:Method2({p:Name})}, {di:Field} ", result.MethodParams);

            result = (MethodCallTemplateParts)method.Invoke(methodCallValueProvider, new[] { "ExcelReporter.Tests.Implementations.Providers:TestClass:Method1( {p:Name}, {m:Method2({p:Name})}, {di:Field} )" });
            Assert.AreEqual("Method1", result.MethodName);
            Assert.AreEqual("ExcelReporter.Tests.Implementations.Providers:TestClass", result.TypeName);
            Assert.AreEqual(" {p:Name}, {m:Method2({p:Name})}, {di:Field} ", result.MethodParams);

            result = (MethodCallTemplateParts)method.Invoke(methodCallValueProvider, new[] { "Method1({p:Name}, {m:Method2({p:Name}, { p:value }, {ms:Method3(hi, 5, {p:Desc})})}, {di:Field})" });
            Assert.AreEqual("Method1", result.MethodName);
            Assert.IsNull(result.TypeName);
            Assert.AreEqual("{p:Name}, {m:Method2({p:Name}, { p:value }, {ms:Method3(hi, 5, {p:Desc})})}, {di:Field}", result.MethodParams);

            result = (MethodCallTemplateParts)method.Invoke(methodCallValueProvider, new[] { "TestClass:Method1({p:Name}, {m:Method2({p:Name}, { p:value }, {ms:Method3(hi, 5, {p:Desc})})}, {di:Field})" });
            Assert.AreEqual("Method1", result.MethodName);
            Assert.AreEqual("TestClass", result.TypeName);
            Assert.AreEqual("{p:Name}, {m:Method2({p:Name}, { p:value }, {ms:Method3(hi, 5, {p:Desc})})}, {di:Field}", result.MethodParams);

            result = (MethodCallTemplateParts)method.Invoke(methodCallValueProvider, new[] { ":TestClass:Method1({p:Name}, {m:Method2({p:Name}, { p:value }, {ms:Method3(hi, 5, {p:Desc})})}, {di:Field})" });
            Assert.AreEqual("Method1", result.MethodName);
            Assert.AreEqual(":TestClass", result.TypeName);
            Assert.AreEqual("{p:Name}, {m:Method2({p:Name}, { p:value }, {ms:Method3(hi, 5, {p:Desc})})}, {di:Field}", result.MethodParams);

            result = (MethodCallTemplateParts)method.Invoke(methodCallValueProvider, new[] { "ExcelReporter.Tests.Implementations.Providers:TestClass:Method1({p:Name}, {m:Method2({p:Name}, { p:value }, {ms:Method3(hi, 5, {p:Desc})})}, {di:Field})" });
            Assert.AreEqual("Method1", result.MethodName);
            Assert.AreEqual("ExcelReporter.Tests.Implementations.Providers:TestClass", result.TypeName);
            Assert.AreEqual("{p:Name}, {m:Method2({p:Name}, { p:value }, {ms:Method3(hi, 5, {p:Desc})})}, {di:Field}", result.MethodParams);
        }

        [TestMethod]
        public void TestParseParams()
        {
            var typeProvider = Substitute.For<ITypeProvider>();
            var methodCallValueProvider = new MethodCallValueProvider(typeProvider);
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

            result = (string[])method.Invoke(methodCallValueProvider, new[] { "{p:Name}" });
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
            MyAssert.Throws<ArgumentNullException>(() => new MethodCallValueProvider(null));

            var typeProvider = Substitute.For<ITypeProvider>();
            typeProvider.GetType(Arg.Any<string>()).Returns(typeof(TestClass));

            var templateProcessor = Substitute.For<ITemplateProcessor>();
            templateProcessor.Pattern.Returns(@"\{.+?:.+?\}");
            templateProcessor.GetValue("{p:Name}").Returns("TestName");
            templateProcessor.GetValue("{p:Desc}").Returns("TestDesc");
            templateProcessor.GetValue("{p:Value}").Returns(7);
            templateProcessor.GetValue("{di.Field}").Returns(777);

            var methodCallValueProvider = new MethodCallValueProvider(typeProvider);
            //Assert.AreEqual("Str", methodCallValueProvider.CallMethod("Method1()", templateProcessor, null));
            //typeProvider.Received(1).GetType(null);
            //templateProcessor.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());

            //typeProvider.ClearReceivedCalls();
            //Assert.AreEqual("Str", methodCallValueProvider.CallMethod("TestClass:Method1()", templateProcessor, null));
            //typeProvider.Received(1).GetType("TestClass");
            //templateProcessor.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());

            //typeProvider.ClearReceivedCalls();
            //Assert.AreEqual("Str", methodCallValueProvider.CallMethod("ExcelReporter.Tests.Implementations.Providers:TestClass:Method1()", templateProcessor, null));
            //typeProvider.Received(1).GetType("ExcelReporter.Tests.Implementations.Providers:TestClass");
            //templateProcessor.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());

            //typeProvider.ClearReceivedCalls();
            //Assert.AreEqual("Str", methodCallValueProvider.CallMethod(":TestClass2:Method1()", templateProcessor, null));
            //typeProvider.Received(1).GetType(":TestClass2");
            //templateProcessor.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());

            typeProvider.ClearReceivedCalls();
            methodCallValueProvider.CallMethod(
                "Method2(5, {p:Name}, hi, { TestClass:Method3({p:Desc}, { ExcelReporter.Tests.Implementations.Providers:TestClass:Method4({str, {di.Field}})} ) }, Method5(), :TestClass2:Method1())",
                templateProcessor, null);
            Assert.AreEqual("Str", methodCallValueProvider.CallMethod(":TestClass2:Method1()", templateProcessor, null));
            typeProvider.Received(1).GetType(":TestClass2");
            templateProcessor.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());

            MyAssert.Throws<ArgumentException>(() => methodCallValueProvider.CallMethod(null, templateProcessor, null));
        }

        private class TestClass
        {
            public string Method1()
            {
                return "Str";
            }

            public int Method2(string arg1, string arg2, string arg3, int arg4, Guid arg5, TestClass2 arg6)
            {
                return 25;
            }

            public static int Method3(int arg1, int arg2)
            {
                return arg1 + arg2;
            }

            public int Method4(string arg1, int arg2)
            {
                return 5;
            }

            public Guid Method5()
            {
                return Guid.NewGuid();
            }
        }
    }
}

public class TestClass2
{
    public TestClass2 Method1()
    {
        return new TestClass2();
    }
}