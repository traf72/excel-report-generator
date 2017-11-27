using System;
using ExcelReporter.Rendering.Providers.ColumnsProviders;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;

namespace ExcelReporter.Tests.Rendering.Providers.ColumnsProvider
{
    [TestClass]
    public class ObjectColumnsProviderTest
    {
        [TestMethod]
        public void TestGetColumnsList()
        {
            IGenericColumnsProvider<Type> typeColumsProvider = Substitute.For<IGenericColumnsProvider<Type>>();

            IColumnsProvider columnsProvider = new ObjectColumnsProvider(typeColumsProvider);
            var testObject = new TypeColumnsProviderTest.TestType();
            columnsProvider.GetColumnsList(testObject);
            typeColumsProvider.Received(1).GetColumnsList(testObject.GetType());

            typeColumsProvider.ClearReceivedCalls();

            var str = "str";
            columnsProvider.GetColumnsList(str);
            typeColumsProvider.Received(1).GetColumnsList(str.GetType());
        }

        [TestMethod]
        public void TestGetColumnsListIfObjectIsNull()
        {
            IColumnsProvider columnsProvider = new ObjectColumnsProvider(new TypeColumnsProvider());
            Assert.AreEqual(0, columnsProvider.GetColumnsList(null).Count);
        }

        [TestMethod]
        public void TestGetColumnsListIfTypeColumnsProviderIsNull()
        {
            ExceptionAssert.Throws<ArgumentNullException>(() => new ObjectColumnsProvider(null));
        }
    }
}