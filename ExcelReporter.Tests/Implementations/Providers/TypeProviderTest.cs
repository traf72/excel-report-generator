using System;
using System.Reflection;
using ExcelReporter.Exceptions;
using ExcelReporter.Implementations.Providers;
using ExcelReporter.Interfaces.Providers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Implementations.Providers
{
    [TestClass]
    public class TypeProviderTest
    {
        [TestMethod]
        public void TestGetType()
        {
            ITypeProvider typeProvider = new TypeProvider(Assembly.GetExecutingAssembly());

            MyAssert.Throws<ArgumentException>(() => typeProvider.GetType(null));
            MyAssert.Throws<ArgumentException>(() => typeProvider.GetType(string.Empty));
            MyAssert.Throws<ArgumentException>(() => typeProvider.GetType(" "));

            Assert.AreSame(typeof(TestType_1), typeProvider.GetType("TestType_1"));
            Assert.AreSame(typeof(TestType_1), typeProvider.GetType("ExcelReporter.Tests.Implementations.Providers:TestType_1"));
            Assert.AreSame(typeof(TestType_1.TestType_2), typeProvider.GetType("TestType_2"));
            Assert.AreSame(typeof(TestType_1.TestType_2), typeProvider.GetType("ExcelReporter.Tests.Implementations.Providers:TestType_2"));
            Assert.AreSame(typeof(TestType_1.TestType_2), typeProvider.GetType(" ExcelReporter.Tests.Implementations.Providers : TestType_2 "));

            Assert.AreSame(typeof(TestType_3), typeProvider.GetType("ExcelReporter.Tests.Implementations.Providers:TestType_3"));
            Assert.AreSame(typeof(InnerNamespace.TestType_3), typeProvider.GetType("ExcelReporter.Tests.Implementations.Providers.InnerNamespace:TestType_3"));
            MyAssert.Throws<IncorrectTemplateException>(() => typeProvider.GetType("TestType_3"), "More than one type found by template \"TestType_3\"");

            Assert.AreSame(typeof(InnerNamespace.TestType_5), typeProvider.GetType("ExcelReporter.Tests.Implementations.Providers.InnerNamespace:TestType_5"));
            Assert.AreSame(typeof(TestType_5), typeProvider.GetType(":TestType_5"));
            MyAssert.Throws<IncorrectTemplateException>(() => typeProvider.GetType("TestType_5"), "More than one type found by template \"TestType_5\"");

            Assert.AreSame(typeof(InnerNamespace.TestType_4), typeProvider.GetType("ExcelReporter.Tests.Implementations.Providers.InnerNamespace:TestType_4"));
            Assert.AreSame(typeof(InnerNamespace.TestType_4), typeProvider.GetType("TestType_4"));
            MyAssert.Throws<IncorrectTemplateException>(() => typeProvider.GetType("ExcelReporter.Tests.Implementations.Providers:TestType_4"),
                "Cannot find type by template \"ExcelReporter.Tests.Implementations.Providers:TestType_4\"");

            MyAssert.Throws<IncorrectTemplateException>(() => typeProvider.GetType("ExcelReporter.Tests.Implementations.Providers:InnerNamespace:TestType_4"),
                "Type name template \"ExcelReporter.Tests.Implementations.Providers:InnerNamespace:TestType_4\" is incorrect");
        }

        private class TestType_1
        {
            public class TestType_2
            {
            }
        }
    }

    public class TestType_3
    {
    }

    namespace InnerNamespace
    {
        public class TestType_3
        {
        }

        public class TestType_4
        {
        }

        public class TestType_5
        {
        }
    }
}

public class TestType_5
{
}