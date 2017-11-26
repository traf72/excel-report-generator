using ExcelReporter.Implementations.Providers.DataItemColumnsProvider;
using ExcelReporter.Interfaces.Providers.DataItemColumnsProvider;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;
using System;

namespace ExcelReporter.Tests.Implementations.Providers.DataItemColumnsProvider
{
    [TestClass]
    public class ObjectColumnsProviderTest
    {
        [TestMethod]
        public void TestGetColumnsList()
        {
            IGenericDataItemColumnsProvider<Type> typeColumsProvider = Substitute.For<IGenericDataItemColumnsProvider<Type>>();

            IDataItemColumnsProvider columnsProvider = new ObjectColumnsProvider(typeColumsProvider);
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
            IDataItemColumnsProvider columnsProvider = new ObjectColumnsProvider(new TypeColumnsProvider());
            Assert.AreEqual(0, columnsProvider.GetColumnsList(null).Count);
        }

        [TestMethod]
        public void TestGetColumnsListIfTypeColumnsProviderIsNull()
        {
            ExceptionAssert.Throws<ArgumentNullException>(() => new ObjectColumnsProvider(null));
        }
    }
}