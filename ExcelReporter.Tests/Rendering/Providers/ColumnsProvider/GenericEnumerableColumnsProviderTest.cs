using System;
using System.Collections.Generic;
using ExcelReporter.Rendering.Providers.ColumnsProviders;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;

namespace ExcelReporter.Tests.Rendering.Providers.ColumnsProvider
{
    [TestClass]
    public class GenericEnumerableColumnsProviderTest
    {
        [TestMethod]
        public void TestGetColumnsList()
        {
            IGenericColumnsProvider<Type> typeColumsProvider = Substitute.For<IGenericColumnsProvider<Type>>();

            IColumnsProvider columnsProvider = new GenericEnumerableColumnsProvider(typeColumsProvider);
            columnsProvider.GetColumnsList(new List<TypeColumnsProviderTest.TestType>());

            typeColumsProvider.Received(1).GetColumnsList(typeof(TypeColumnsProviderTest.TestType));

            typeColumsProvider.ClearReceivedCalls();

            columnsProvider.GetColumnsList(new List<string>());
            typeColumsProvider.Received(1).GetColumnsList(typeof(string));
        }

        [TestMethod]
        public void TestGetColumnsListIfEnumerableIsNull()
        {
            IColumnsProvider columnsProvider = new GenericEnumerableColumnsProvider(new TypeColumnsProvider());
            Assert.AreEqual(0, columnsProvider.GetColumnsList(null).Count);
        }

        [TestMethod]
        public void TestGetColumnsListIfTypeColumnsProviderIsNull()
        {
            ExceptionAssert.Throws<ArgumentNullException>(() => new GenericEnumerableColumnsProvider(null));
        }
    }
}