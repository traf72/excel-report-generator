using ExcelReportGenerator.Excel;
using ExcelReportGenerator.Exceptions;
using ExcelReportGenerator.Rendering.Providers;
using ExcelReportGenerator.Tests.CustomAsserts;
using NUnit.Framework;
using System;
using System.Reflection;

namespace ExcelReportGenerator.Tests.Rendering.Providers
{
    
    public class DefaulTypeProviderTest
    {
        [Test]
        public void TestGetType()
        {
            //// These tests do not pass because the entry assembly exists after migration to .NET Core
            //ExceptionAssert.Throws<InvalidOperationException>(() => new DefaultTypeProvider(), "Assemblies are not provided but entry assembly is null. Provide assemblies and try again.");
            //ExceptionAssert.Throws<InvalidOperationException>(() => new DefaultTypeProvider(new Assembly[0]), "Assemblies are not provided but entry assembly is null. Provide assemblies and try again.");

            var typeProvider = new DefaultTypeProvider(new[] { typeof(ExcelHelper).Assembly });
            Assert.AreSame(typeof(ExcelHelper), typeProvider.GetType("ExcelHelper"));
            ExceptionAssert.Throws<TypeNotFoundException>(() => typeProvider.GetType("TestType_1"), "Cannot find type by template \"TestType_1\"");
            ExceptionAssert.Throws<TypeNotFoundException>(() => typeProvider.GetType("DateTime"), "Cannot find type by template \"DateTime\"");

            ExceptionAssert.Throws<InvalidOperationException>(() => typeProvider.GetType(null), "Template is not specified but defaultType is null");
            ExceptionAssert.Throws<InvalidOperationException>(() => typeProvider.GetType(string.Empty), "Template is not specified but defaultType is null");
            ExceptionAssert.Throws<InvalidOperationException>(() => typeProvider.GetType(" "), "Template is not specified but defaultType is null");

            typeProvider = new DefaultTypeProvider(new[] { Assembly.GetExecutingAssembly() }, typeof(TestType_1));

            Assert.AreSame(typeof(TestType_1), typeProvider.GetType("TestType_1"));
            Assert.AreSame(typeof(TestType_1), typeProvider.GetType(null));
            Assert.AreSame(typeProvider.GetType("TestType_1"), typeProvider.GetType(null));
            Assert.AreSame(typeof(TestType_1), typeProvider.GetType("ExcelReportGenerator.Tests.Rendering.Providers:TestType_1"));
            Assert.AreSame(typeof(TestType_1.TestType_2), typeProvider.GetType("TestType_2"));
            Assert.AreSame(typeof(TestType_1.TestType_2), typeProvider.GetType("ExcelReportGenerator.Tests.Rendering.Providers:TestType_2"));
            Assert.AreSame(typeof(TestType_1.TestType_2), typeProvider.GetType(" ExcelReportGenerator.Tests.Rendering.Providers : TestType_2 "));
            ExceptionAssert.Throws<TypeNotFoundException>(() => typeProvider.GetType("ExcelHelper"), "Cannot find type by template \"ExcelHelper\"");

            Assert.AreSame(typeof(TestType_3), typeProvider.GetType("ExcelReportGenerator.Tests.Rendering.Providers:TestType_3"));
            Assert.AreSame(typeof(InnerNamespace.TestType_3), typeProvider.GetType("ExcelReportGenerator.Tests.Rendering.Providers.InnerNamespace:TestType_3"));
            ExceptionAssert.Throws<InvalidTemplateException>(() => typeProvider.GetType("TestType_3"), "More than one type found by template \"TestType_3\"");
            ExceptionAssert.Throws<TypeNotFoundException>(() => typeProvider.GetType("DateTime"), "Cannot find type by template \"DateTime\"");

            typeProvider = new DefaultTypeProvider(new[] { Assembly.GetExecutingAssembly(), Assembly.GetAssembly(typeof(DateTime)) });

            Assert.AreSame(typeof(InnerNamespace.TestType_5), typeProvider.GetType("ExcelReportGenerator.Tests.Rendering.Providers.InnerNamespace:TestType_5"));
            Assert.AreSame(typeof(TestType_5), typeProvider.GetType(":TestType_5"));
            Assert.AreSame(typeof(TestType_5), typeProvider.GetType(":TestType_5"));
            Assert.AreSame(typeof(DateTime), typeProvider.GetType("DateTime"));
            Assert.AreSame(typeof(DateTime), typeProvider.GetType("System:DateTime"));
            ExceptionAssert.Throws<InvalidTemplateException>(() => typeProvider.GetType("TestType_5"), "More than one type found by template \"TestType_5\"");

            Assert.AreSame(typeof(InnerNamespace.TestType_4), typeProvider.GetType("ExcelReportGenerator.Tests.Rendering.Providers.InnerNamespace:TestType_4"));
            Assert.AreSame(typeof(InnerNamespace.TestType_4), typeProvider.GetType("TestType_4"));
            ExceptionAssert.Throws<TypeNotFoundException>(() => typeProvider.GetType("ExcelReportGenerator.Tests.Rendering.Providers:TestType_4"),
                "Cannot find type by template \"ExcelReportGenerator.Tests.Rendering.Providers:TestType_4\"");

            ExceptionAssert.Throws<InvalidTemplateException>(() => typeProvider.GetType("ExcelReportGenerator.Tests.Rendering.Providers:InnerNamespace:TestType_4"),
                "Type name template \"ExcelReportGenerator.Tests.Rendering.Providers:InnerNamespace:TestType_4\" is invalid");
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