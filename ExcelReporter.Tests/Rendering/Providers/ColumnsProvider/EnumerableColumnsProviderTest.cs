using System;
using System.Collections;
using ExcelReporter.Rendering.Providers.ColumnsProviders;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;

namespace ExcelReporter.Tests.Rendering.Providers.ColumnsProvider
{
    [TestClass]
    public class EnumerableColumnsProviderTest
    {
        [TestMethod]
        public void TestGetColumnsList()
        {
            IGenericDataItemColumnsProvider<Type> typeColumsProvider = Substitute.For<IGenericDataItemColumnsProvider<Type>>();
            IDataItemColumnsProvider columnsProvider = new EnumerableColumnsProvider(typeColumsProvider);

            columnsProvider.GetColumnsList(columnsProvider.GetColumnsList(new ArrayList { new TypeColumnsProviderTest.TestType(), "str" }));
            typeColumsProvider.Received(1).GetColumnsList(typeof(TypeColumnsProviderTest.TestType));

            typeColumsProvider.ClearReceivedCalls();

            columnsProvider.GetColumnsList(columnsProvider.GetColumnsList(new ArrayList { "str", new TypeColumnsProviderTest.TestType() }));
            typeColumsProvider.Received(1).GetColumnsList(typeof(string));
        }

        [TestMethod]
        public void TestGetColumnsListIfEnumerableIsNullOrEmpty()
        {
            IDataItemColumnsProvider columnsProvider = new EnumerableColumnsProvider(new TypeColumnsProvider());
            Assert.AreEqual(0, columnsProvider.GetColumnsList(null).Count);
            Assert.AreEqual(0, columnsProvider.GetColumnsList(new ArrayList()).Count);
        }

        [TestMethod]
        public void TestGetColumnsListIfTypeColumnsProviderIsNull()
        {
            ExceptionAssert.Throws<ArgumentNullException>(() => new EnumerableColumnsProvider(null));
        }
    }
}