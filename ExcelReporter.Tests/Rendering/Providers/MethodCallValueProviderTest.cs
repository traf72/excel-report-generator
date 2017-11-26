using System;
using System.Reflection;
using ExcelReporter.Exceptions;
using ExcelReporter.Rendering;
using ExcelReporter.Rendering.Providers;
using ExcelReporter.Rendering.TemplateProcessors;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;

namespace ExcelReporter.Tests.Rendering.Providers
{
    [TestClass]
    public class MethodCallValueProviderTest
    {
        [TestMethod]
        public void TestParseTemplate()
        {
            var typeProvider = Substitute.For<ITypeProvider>();
            var methodCallValueProvider = new DefaultMethodCallValueProvider(typeProvider, null);
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

            result = method.Invoke(methodCallValueProvider, new[] { "Method1(,)" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.IsNull(typeNameProp.GetValue(result));
            Assert.AreEqual(",", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "Method1( , )" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.IsNull(typeNameProp.GetValue(result));
            Assert.AreEqual(",", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "Method1(\"\",\" \")" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.IsNull(typeNameProp.GetValue(result));
            Assert.AreEqual("\"\",\" \"", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "Method1(()" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.IsNull(typeNameProp.GetValue(result));
            Assert.AreEqual("(", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "Method1(m:Method2(p:Name))" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.IsNull(typeNameProp.GetValue(result));
            Assert.AreEqual("m:Method2(p:Name)", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "TestClass:Method1(m:Method2(p:Name))" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.AreEqual("TestClass", typeNameProp.GetValue(result));
            Assert.AreEqual("m:Method2(p:Name)", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { ":TestClass:Method1(m:Method2(p:Name))" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.AreEqual(":TestClass", typeNameProp.GetValue(result));
            Assert.AreEqual("m:Method2(p:Name)", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "ExcelReporter.Tests.Implementations.Providers:TestClass:Method1(m:Method2(p:Name))" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.AreEqual("ExcelReporter.Tests.Implementations.Providers:TestClass", typeNameProp.GetValue(result));
            Assert.AreEqual("m:Method2(p:Name)", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "Method1( p:Name, m:Method2(p:Name), di:Field )" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.IsNull(typeNameProp.GetValue(result));
            Assert.AreEqual("p:Name, m:Method2(p:Name), di:Field", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "TestClass:Method1( p:Name, m:Method2(p:Name), di:Field )" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.AreEqual("TestClass", typeNameProp.GetValue(result));
            Assert.AreEqual("p:Name, m:Method2(p:Name), di:Field", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { ":TestClass:Method1( p:Name, m:Method2(p:Name), di:Field )" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.AreEqual(":TestClass", typeNameProp.GetValue(result));
            Assert.AreEqual("p:Name, m:Method2(p:Name), di:Field", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "ExcelReporter.Tests.Implementations.Providers:TestClass:Method1( p:Name, m:Method2(p:Name), di:Field )" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.AreEqual("ExcelReporter.Tests.Implementations.Providers:TestClass", typeNameProp.GetValue(result));
            Assert.AreEqual("p:Name, m:Method2(p:Name), di:Field", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "Method1(p:Name, m:Method2(p:Name,  p:value , ms:Method3(\"hi\", 5, p:Desc)), di:Field)" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.IsNull(typeNameProp.GetValue(result));
            Assert.AreEqual("p:Name, m:Method2(p:Name,  p:value , ms:Method3(\"hi\", 5, p:Desc)), di:Field", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "TestClass:Method1(p:Name, m:Method2(p:Name,  p:value , ms:Method3(\"(\", 5, p:Desc)), di:Field)" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.AreEqual("TestClass", typeNameProp.GetValue(result));
            Assert.AreEqual("p:Name, m:Method2(p:Name,  p:value , ms:Method3(\"(\", 5, p:Desc)), di:Field", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { ":TestClass:Method1(p:Name, m:Method2(p:Name,  p:value , ms:Method3(hi, 5, p:Desc)), di:Field)" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.AreEqual(":TestClass", typeNameProp.GetValue(result));
            Assert.AreEqual("p:Name, m:Method2(p:Name,  p:value , ms:Method3(hi, 5, p:Desc)), di:Field", methodParamsProp.GetValue(result));

            result = method.Invoke(methodCallValueProvider, new[] { "ExcelReporter.Tests.Implementations.Providers:TestClass:Method1(p:Name, m:Method2(p:Name,  p:value , ms:Method3(hi, 5, p:Desc)), di:Field)" });
            Assert.AreEqual("Method1", methodNameProp.GetValue(result));
            Assert.AreEqual("ExcelReporter.Tests.Implementations.Providers:TestClass", typeNameProp.GetValue(result));
            Assert.AreEqual("p:Name, m:Method2(p:Name,  p:value , ms:Method3(hi, 5, p:Desc)), di:Field", methodParamsProp.GetValue(result));

            ExceptionAssert.ThrowsBaseException<IncorrectTemplateException>(() => method.Invoke(methodCallValueProvider, new[] { "Method1" }), "Template \"Method1\" is incorrect");
            ExceptionAssert.ThrowsBaseException<IncorrectTemplateException>(() => method.Invoke(methodCallValueProvider, new[] { "Method1(" }), "Template \"Method1(\" is incorrect");
            ExceptionAssert.ThrowsBaseException<IncorrectTemplateException>(() => method.Invoke(methodCallValueProvider, new[] { "Method1)" }), "Template \"Method1)\" is incorrect");
        }

        [TestMethod]
        public void TestParseParams()
        {
            var typeProvider = Substitute.For<ITypeProvider>();
            var methodCallValueProvider = new DefaultMethodCallValueProvider(typeProvider, null);
            MethodInfo method = methodCallValueProvider.GetType().GetMethod("ParseInputParams", BindingFlags.Instance | BindingFlags.NonPublic);

            var result = (string[])method.Invoke(methodCallValueProvider, new[] { string.Empty });
            Assert.AreEqual(0, result.Length);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { (string)null });
            Assert.AreEqual(0, result.Length);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { "\"\"" });
            Assert.AreEqual(1, result.Length);
            Assert.AreEqual("\"\"", result[0]);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { " " });
            Assert.AreEqual(0, result.Length);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { "\" \"" });
            Assert.AreEqual(1, result.Length);
            Assert.AreEqual("\" \"", result[0]);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { "4" });
            Assert.AreEqual(1, result.Length);
            Assert.AreEqual("4", result[0]);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { " p:Name " });
            Assert.AreEqual(1, result.Length);
            Assert.AreEqual("p:Name", result[0]);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { " m:Method() " });
            Assert.AreEqual(1, result.Length);
            Assert.AreEqual("m:Method()", result[0]);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { "ms:TestClass:Method()" });
            Assert.AreEqual(1, result.Length);
            Assert.AreEqual("ms:TestClass:Method()", result[0]);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { "m:Namespace.TestClass:Method()" });
            Assert.AreEqual(1, result.Length);
            Assert.AreEqual("m:Namespace.TestClass:Method()", result[0]);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { "m::TestClass:Method()" });
            Assert.AreEqual(1, result.Length);
            Assert.AreEqual("m::TestClass:Method()", result[0]);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { "ms:Method(p:Name)" });
            Assert.AreEqual(1, result.Length);
            Assert.AreEqual("ms:Method(p:Name)", result[0]);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { "p:Name, 11, m:TestClass:Method(p:Name), di:Desc, [string]hi, m:Namespace.TestClass:Method( ms:Method2() )" });
            Assert.AreEqual(6, result.Length);
            Assert.AreEqual("p:Name", result[0]);
            Assert.AreEqual("11", result[1]);
            Assert.AreEqual("m:TestClass:Method(p:Name)", result[2]);
            Assert.AreEqual("di:Desc", result[3]);
            Assert.AreEqual("[string]hi", result[4]);
            Assert.AreEqual("m:Namespace.TestClass:Method( ms:Method2() )", result[5]);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { " m:Method(p:Name, di:Desc) " });
            Assert.AreEqual(1, result.Length);
            Assert.AreEqual("m:Method(p:Name, di:Desc)", result[0]);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { " m:Method(p:Name, di:Desc) , [Int32] 11,  m:Method2(m:Namespace.TestClass:Method( ms:Method2(di:Field, [long]777) , 12, p:Name), [short]11, p:Value) , ," });
            Assert.AreEqual(5, result.Length);
            Assert.AreEqual("m:Method(p:Name, di:Desc)", result[0]);
            Assert.AreEqual("[Int32] 11", result[1]);
            Assert.AreEqual("m:Method2(m:Namespace.TestClass:Method( ms:Method2(di:Field, [long]777) , 12, p:Name), [short]11, p:Value)", result[2]);
            Assert.AreEqual(string.Empty, result[3]);
            Assert.AreEqual(string.Empty, result[4]);

            // Экранирование
            result = (string[])method.Invoke(methodCallValueProvider, new[] { "Привет,, медвед!,m:Method(10,p:Name)" });
            Assert.AreEqual(2, result.Length);
            Assert.AreEqual("Привет, медвед!", result[0]);
            Assert.AreEqual("m:Method(10,p:Name)", result[1]);

            result = (string[])method.Invoke(methodCallValueProvider, new[] { "(Test1), \"(Test2)\", (), m:Method((Test3), \"(Test4)\", (), (Test5))" });
            Assert.AreEqual(4, result.Length);
            Assert.AreEqual("(Test1)", result[0]);
            Assert.AreEqual("\"(Test2)\"", result[1]);
            Assert.AreEqual("()", result[2]);
            Assert.AreEqual("m:Method((Test3), \"(Test4)\", (), (Test5))", result[3]);
        }

        [TestMethod]
        public void TestCallMethod()
        {
            ExceptionAssert.Throws<ArgumentNullException>(() => new DefaultMethodCallValueProvider(null, new object()));

            var typeProvider = Substitute.For<ITypeProvider>();
            typeProvider.GetType(Arg.Any<string>()).Returns(typeof(TestClass));

            var templateProcessor = Substitute.For<ITemplateProcessor>();
            var dataItem = new HierarchicalDataItem();
            templateProcessor.TemplatePattern.Returns(@"{.+?:.+?}");
            templateProcessor.LeftTemplateBorder.Returns("{");
            templateProcessor.RightTemplateBorder.Returns("}");
            templateProcessor.GetValue("p:Name").Returns("TestName");
            templateProcessor.GetValue("p:Value", dataItem).Returns(7);

            var methodCallValueProvider = new DefaultMethodCallValueProvider(typeProvider, new TestClass());

            ExceptionAssert.Throws<ArgumentException>(() => methodCallValueProvider.CallMethod(null, templateProcessor, new HierarchicalDataItem()));
            ExceptionAssert.Throws<ArgumentException>(() => methodCallValueProvider.CallMethod(string.Empty, templateProcessor, new HierarchicalDataItem()));
            ExceptionAssert.Throws<ArgumentException>(() => methodCallValueProvider.CallMethod(" ", templateProcessor, new HierarchicalDataItem()));

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
            Assert.AreEqual(25, methodCallValueProvider.CallMethod("Method2(p:Value, 18)", templateProcessor, dataItem, true));
            typeProvider.DidNotReceiveWithAnyArgs().GetType(Arg.Any<string>());
            templateProcessor.Received(1).GetValue("p:Value", dataItem);

            typeProvider.ClearReceivedCalls();
            templateProcessor.ClearReceivedCalls();
            Assert.AreEqual(25, methodCallValueProvider.CallMethod(" : TestClass : Method2(p:Value, 18) ", templateProcessor, dataItem, true));
            typeProvider.Received(1).GetType(": TestClass");
            templateProcessor.Received(1).GetValue("p:Value", dataItem);

            typeProvider.ClearReceivedCalls();
            templateProcessor.ClearReceivedCalls();
            Assert.IsNull(methodCallValueProvider.CallMethod("Method3()", templateProcessor, null));
            typeProvider.DidNotReceiveWithAnyArgs().GetType(Arg.Any<string>());
            templateProcessor.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());

            templateProcessor.GetValue("m:TestClass:Method5(p:Desc,  ms:ExcelReporter.Tests.Implementations.Providers:TestClass:Method6(str, di.Field) )").Returns(10);
            templateProcessor.GetValue("m:Method7()").Returns('c');
            templateProcessor.GetValue("m::TestClass2:Method1()").Returns(long.MaxValue);

            object result = methodCallValueProvider.CallMethod(
                "Method4(5, p:Name, hi,  m:TestClass:Method5(p:Desc,  ms:ExcelReporter.Tests.Implementations.Providers:TestClass:Method6(str, di.Field) ) , m:Method7(), m::TestClass2:Method1())",
                templateProcessor, null);
            Assert.AreEqual($"5_TestName_hi_10_c_{long.MaxValue}", result);
            typeProvider.DidNotReceiveWithAnyArgs().GetType(Arg.Any<string>());
            templateProcessor.Received(1).GetValue("p:Name");
            templateProcessor.Received(1).GetValue("m:TestClass:Method5(p:Desc,  ms:ExcelReporter.Tests.Implementations.Providers:TestClass:Method6(str, di.Field) )");
            templateProcessor.Received(1).GetValue("m:Method7()");
            templateProcessor.Received(1).GetValue("m::TestClass2:Method1()");

            templateProcessor.ClearReceivedCalls();
            Assert.AreEqual("_ ", methodCallValueProvider.CallMethod("Method5(\"\", \" \")", templateProcessor, null));
            typeProvider.DidNotReceiveWithAnyArgs().GetType(Arg.Any<string>());
            templateProcessor.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());

            Assert.AreEqual("p:Name_m:Method6()", methodCallValueProvider.CallMethod("Method5(\"p:Name\", \"m:Method6()\")", templateProcessor, null));
            typeProvider.DidNotReceiveWithAnyArgs().GetType(Arg.Any<string>());
            templateProcessor.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());

            Assert.AreEqual("\"p:Name\"_\"\"", methodCallValueProvider.CallMethod("Method5(\"\"p:Name\"\", \"\"\"\")", templateProcessor, null));
            typeProvider.DidNotReceiveWithAnyArgs().GetType(Arg.Any<string>());
            templateProcessor.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());

            ExceptionAssert.Throws<MethodNotFoundException>(() => methodCallValueProvider.CallMethod("TestClass:BadMethod()", templateProcessor, null),
                "Could not find public method \"BadMethod\" in type \"TestClass\" and all its parents. MethodCallTemplate: TestClass:BadMethod()");

            ExceptionAssert.Throws<MethodNotFoundException>(() => methodCallValueProvider.CallMethod("TestClass:BadMethod()", templateProcessor, null, true),
                "Could not find public static method \"BadMethod\" in type \"TestClass\" and all its parents. MethodCallTemplate: TestClass:BadMethod()");

            typeProvider.ClearReceivedCalls();
            templateProcessor.ClearReceivedCalls();
            Assert.AreEqual("Str_Parent", methodCallValueProvider.CallMethod("MethodParent()", templateProcessor, null));
            typeProvider.DidNotReceiveWithAnyArgs().GetType(Arg.Any<string>());
            templateProcessor.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());

            Assert.AreEqual("Str_Static_Parent", methodCallValueProvider.CallMethod("MethodStaticParent()", templateProcessor, null, true));
            typeProvider.DidNotReceiveWithAnyArgs().GetType(Arg.Any<string>());
            templateProcessor.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());

            methodCallValueProvider = new DefaultMethodCallValueProvider(typeProvider, null);

            Assert.AreEqual("Str_1", methodCallValueProvider.CallMethod(":TestClass:Method1()", templateProcessor, null));
            typeProvider.Received(1).GetType(":TestClass");
            templateProcessor.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());

            ExceptionAssert.Throws<InvalidOperationException>(() => methodCallValueProvider.CallMethod("Method2(p:Value, 18)", templateProcessor, null),
                "Type name is not specified in template \"Method2(p:Value, 18)\" but defaultInstance is null");
        }

        [TestMethod]
        public void TestCallMethodWithOverloading()
        {
            var typeProvider = Substitute.For<ITypeProvider>();
            typeProvider.GetType(Arg.Any<string>()).Returns(typeof(TestOverloading));

            var templateProcessor = Substitute.For<ITemplateProcessor>();
            templateProcessor.TemplatePattern.Returns(@"{.+?:.+?}");
            templateProcessor.LeftTemplateBorder.Returns("{");
            templateProcessor.RightTemplateBorder.Returns("}");
            templateProcessor.GetValue("p:Name").Returns("TestName");
            templateProcessor.GetValue("p:Value").Returns(7);
            templateProcessor.GetValue("p:Value2").Returns((short)77);
            templateProcessor.GetValue("p:Value3").Returns(null);

            var methodCallValueProvider = new DefaultMethodCallValueProvider(typeProvider, new TestOverloading());

            Assert.AreEqual("Method1()", methodCallValueProvider.CallMethod("Method1()", templateProcessor, null));

            Assert.AreEqual("Method2(int), a = 15", methodCallValueProvider.CallMethod("Method2([int]15)", templateProcessor, null));
            Assert.AreEqual("Method2(int), a = 15", methodCallValueProvider.CallMethod("Method2([ Int32 ] 15)", templateProcessor, null));
            Assert.AreEqual("Method2(int), a = 7", methodCallValueProvider.CallMethod("Method2(p:Value)", templateProcessor, null));
            Assert.AreEqual("Method2(string), a = str", methodCallValueProvider.CallMethod("Method2([string]str)", templateProcessor, null));
            Assert.AreEqual("Method2(string), a = str", methodCallValueProvider.CallMethod("Method2(\"str\")", templateProcessor, null));
            ExceptionAssert.Throws<NotSupportedException>(() => methodCallValueProvider.CallMethod("Method2(15)", templateProcessor, null), "More than one method found with suitable number of parameters but some of static parameters does not specify a type explicitly. Specify the type explicitly for all static parameters and try again. MethodCallTemplate: Method2(15)");
            ExceptionAssert.Throws<NotSupportedException>(() => methodCallValueProvider.CallMethod("Method2([short] 15)", templateProcessor, null), "More than one method found with suitable number of parameters. In this case the method is chosen by exact match of parameter types. None of methods is suitable. MethodCallTemplate: Method2([short] 15)");
            ExceptionAssert.Throws<NotSupportedException>(() => methodCallValueProvider.CallMethod("Method2(p:Value2)", templateProcessor, null), "More than one method found with suitable number of parameters. In this case the method is chosen by exact match of parameter types. None of methods is suitable. MethodCallTemplate: Method2(p:Value2)");

            Assert.AreEqual("Method2(int, string), a = 15, b = str", methodCallValueProvider.CallMethod("Method2([int]15, \"str\")", templateProcessor, null));
            Assert.AreEqual("Method2(int, string), a = 15, b = str", methodCallValueProvider.CallMethod("Method2([Int32]15, [String]str)", templateProcessor, null));
            Assert.AreEqual("Method2(int, string), a = 15, b = TestName", methodCallValueProvider.CallMethod("Method2([Int32]15, p:Name)", templateProcessor, null));
            Assert.AreEqual("Method2(int, string), a = 7, b = TestName", methodCallValueProvider.CallMethod("Method2(p:Value, p:Name)", templateProcessor, null));
            Assert.AreEqual("Method2(int, string), a = 7, b = p:Name", methodCallValueProvider.CallMethod("Method2(p:Value, \"p:Name\")", templateProcessor, null));
            ExceptionAssert.Throws<NotSupportedException>(() => methodCallValueProvider.CallMethod("Method2(15, str)", templateProcessor, null), "More than one method found with suitable number of parameters but some of static parameters does not specify a type explicitly. Specify the type explicitly for all static parameters and try again. MethodCallTemplate: Method2(15, str)");
            ExceptionAssert.Throws<NotSupportedException>(() => methodCallValueProvider.CallMethod("Method2([int]15, str)", templateProcessor, null), "More than one method found with suitable number of parameters but some of static parameters does not specify a type explicitly. Specify the type explicitly for all static parameters and try again. MethodCallTemplate: Method2([int]15, str)");
            ExceptionAssert.Throws<NotSupportedException>(() => methodCallValueProvider.CallMethod("Method2(15, [string]str)", templateProcessor, null), "More than one method found with suitable number of parameters but some of static parameters does not specify a type explicitly. Specify the type explicitly for all static parameters and try again. MethodCallTemplate: Method2(15, [string]str)");
            ExceptionAssert.Throws<NotSupportedException>(() => methodCallValueProvider.CallMethod("Method2(15, \"str\")", templateProcessor, null), "More than one method found with suitable number of parameters but some of static parameters does not specify a type explicitly. Specify the type explicitly for all static parameters and try again. MethodCallTemplate: Method2(15, \"str\")");
            ExceptionAssert.Throws<NotSupportedException>(() => methodCallValueProvider.CallMethod("Method2(p:Value2, p:Name)", templateProcessor, null), "More than one method found with suitable number of parameters. In this case the method is chosen by exact match of parameter types. None of methods is suitable. MethodCallTemplate: Method2(p:Value2, p:Name)");

            Assert.AreEqual("Method2(string, int), a = str, b = 15", methodCallValueProvider.CallMethod("Method2(\"str\", [int]15)", templateProcessor, null));
            Assert.AreEqual("Method2(string, int), a = str, b = 15", methodCallValueProvider.CallMethod("Method2([String]str, [Int32]15)", templateProcessor, null));
            Assert.AreEqual("Method2(string, int), a = TestName, b = 7", methodCallValueProvider.CallMethod("Method2(p:Name, p:Value)", templateProcessor, null));

            Assert.AreEqual("Method2(int, string, long), a = 15, b = str, c = 20", methodCallValueProvider.CallMethod("Method2([int]15, [string]str, [long]20)", templateProcessor, null));
            ExceptionAssert.Throws<NotSupportedException>(() => methodCallValueProvider.CallMethod("Method2([int]15, [string]str, 20)", templateProcessor, null), "More than one method found with suitable number of parameters but some of static parameters does not specify a type explicitly. Specify the type explicitly for all static parameters and try again. MethodCallTemplate: Method2([int]15, [string]str, 20)");
            ExceptionAssert.Throws<NotSupportedException>(() => methodCallValueProvider.CallMethod("Method2([int]15, [string]str, [short]20)", templateProcessor, null), "More than one method found with suitable number of parameters. In this case the method is chosen by exact match of parameter types. None of methods is suitable. MethodCallTemplate: Method2([int]15, [string]str, [short]20)");

            Assert.AreEqual("Method2(int, string, long, short), a = 15, b = str, c = 20, d = 200", methodCallValueProvider.CallMethod("Method2(15, str, 20, 200)", templateProcessor, null));
            Assert.AreEqual("Method2(int, string, long, short), a = 15, b = str, c = 20, d = 200", methodCallValueProvider.CallMethod("Method2([short]15, str, 20, 200)", templateProcessor, null));
            Assert.AreEqual("Method2(int, string, long, short), a = 7, b = TestName, c = 7, d = 77", methodCallValueProvider.CallMethod("Method2(p:Value, p:Name, p:Value, p:Value2)", templateProcessor, null));
            ExceptionAssert.Throws<ArgumentException>(() => methodCallValueProvider.CallMethod("Method2(p:Value, p:Name, p:Value2, p:Value)", templateProcessor, null));
            ExceptionAssert.Throws<FormatException>(() => methodCallValueProvider.CallMethod("Method2(p:Value, p:Name, str, p:Value)", templateProcessor, null));
            ExceptionAssert.Throws<MethodNotFoundException>(() => methodCallValueProvider.CallMethod("Method2(15, str, 20, 200, str2)", templateProcessor, null), "Could not find public method \"Method2\" in type \"TestOverloading\" and all its parents with suitable number of parameters. MethodCallTemplate: Method2(15, str, 20, 200, str2)");

            Assert.AreEqual("Method3(int, string, sbyte), a = 15, b = str, c = 1", methodCallValueProvider.CallMethod("Method3(15)", templateProcessor, null));
            Assert.AreEqual("Method3(int, string, sbyte), a = 15, b = str2, c = 1", methodCallValueProvider.CallMethod("Method3(15, str2)", templateProcessor, null));
            Assert.AreEqual("Method3(int, string, sbyte), a = 15, b = str2, c = 127", methodCallValueProvider.CallMethod("Method3(15, str2, 127)", templateProcessor, null));
            Assert.AreEqual("Method3(int, string, sbyte), a = 15, b = str2, c = 0", methodCallValueProvider.CallMethod("Method3(15, str2, p:Value3)", templateProcessor, null));
            Assert.AreEqual("Method3(int, string, sbyte), a = 0, b = str2, c = 1", methodCallValueProvider.CallMethod("Method3(p:Value3, str2)", templateProcessor, null));
            ExceptionAssert.Throws<FormatException>(() => methodCallValueProvider.CallMethod("Method3(str, str2, 127)", templateProcessor, null));
            ExceptionAssert.Throws<FormatException>(() => methodCallValueProvider.CallMethod("Method3([int]str, str2, 127)", templateProcessor, null));
            ExceptionAssert.Throws<NotSupportedException>(() => methodCallValueProvider.CallMethod("Method3([int33]15)", templateProcessor, null), "Type \"int33\" is not supported");
            ExceptionAssert.Throws<OverflowException>(() => methodCallValueProvider.CallMethod("Method3(15, str2, 200)", templateProcessor, null));
            ExceptionAssert.Throws<InvalidOperationException>(() => methodCallValueProvider.CallMethod("Method3()", templateProcessor, null), "Mismatch parameters count. Input pararameters count: 0. Method required parameters count: 1. MethodCallTemplate: Method3()");
            ExceptionAssert.Throws<InvalidOperationException>(() => methodCallValueProvider.CallMethod("Method3(1, 2, 3, 4)", templateProcessor, null), "Mismatch parameters count. Input pararameters count: 4. Method parameters count: 3. MethodCallTemplate: Method3(1, 2, 3, 4)");

            ExceptionAssert.Throws<NotSupportedException>(() => methodCallValueProvider.CallMethod("Method4([int]15)", templateProcessor, null), "Methods which have \"params\" argument are not supported. MethodCallTemplate: Method4([int]15)");
            ExceptionAssert.Throws<MethodNotFoundException>(() => methodCallValueProvider.CallMethod("Method5()", templateProcessor, null), "Could not find public method \"Method5\" in type \"TestOverloading\" and all its parents. MethodCallTemplate: Method5()");
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

            public string Method5(string arg1, string arg2)
            {
                return $"{arg1}_{arg2}";
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

        private class TestOverloading : TestOverloadingParent
        {
            public string Method1()
            {
                return "Method1()";
            }

            public string Method2(int a)
            {
                return $"Method2(int), a = {a}";
            }

            public string Method2(int a, string b)
            {
                return $"Method2(int, string), a = {a}, b = {b}";
            }

            public string Method2(int a, string b, long c)
            {
                return $"Method2(int, string, long), a = {a}, b = {b}, c = {c}";
            }

            public string Method2(string a)
            {
                return $"Method2(string), a = {a}";
            }

            public string Method2(string a, int b = 10)
            {
                return $"Method2(string, int), a = {a}, b = {b}";
            }

            internal string Method2(int a, string b, short c)
            {
                return $"Method2(int, string, short), a = {a}, b = {b}, c = {c}";
            }

            public string Method3(int a, string b = "str", sbyte c = 1)
            {
                return $"Method3(int, string, sbyte), a = {a}, b = {b}, c = {c}";
            }

            public string Method4(int a, params string[] b)
            {
                return $"Method4(int, params string[]), a = {a}, b = {b}";
            }
        }

        private class TestOverloadingParent
        {
            public string Method2(int a, string b, long c, short d = 16)
            {
                return $"Method2(int, string, long, short), a = {a}, b = {b}, c = {c}, d = {d}";
            }
        }
    }
}