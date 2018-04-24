using System;
using System.Collections;
using ExcelReportGenerator.Rendering.Providers.ColumnsProviders;
using ExcelReportGenerator.Tests.CustomAsserts;
using NUnit.Framework;
using NSubstitute;

namespace ExcelReportGenerator.Tests.Rendering.Providers.ColumnsProvider
{
    
    public class EnumerableColumnsProviderTest
    {
        [Test]
        public void TestGetColumnsList()
        {
            IGenericColumnsProvider<Type> typeColumsProvider = Substitute.For<IGenericColumnsProvider<Type>>();
            IColumnsProvider columnsProvider = new EnumerableColumnsProvider(typeColumsProvider);

            columnsProvider.GetColumnsList(columnsProvider.GetColumnsList(new ArrayList { new TypeColumnsProviderTest.TestType(), "str" }));
            typeColumsProvider.Received(1).GetColumnsList(typeof(TypeColumnsProviderTest.TestType));

            typeColumsProvider.ClearReceivedCalls();

            columnsProvider.GetColumnsList(columnsProvider.GetColumnsList(new ArrayList { "str", new TypeColumnsProviderTest.TestType() }));
            typeColumsProvider.Received(1).GetColumnsList(typeof(string));
        }

        [Test]
        public void TestGetColumnsListIfEnumerableIsNullOrEmpty()
        {
            IColumnsProvider columnsProvider = new EnumerableColumnsProvider(new TypeColumnsProvider());
            Assert.AreEqual(0, columnsProvider.GetColumnsList(null).Count);
            Assert.AreEqual(0, columnsProvider.GetColumnsList(new ArrayList()).Count);
        }

        [Test]
        public void TestGetColumnsListIfTypeColumnsProviderIsNull()
        {
            ExceptionAssert.Throws<ArgumentNullException>(() => new EnumerableColumnsProvider(null));
        }
    }
}