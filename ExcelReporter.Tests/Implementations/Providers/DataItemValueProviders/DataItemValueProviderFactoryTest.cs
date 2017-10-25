using ExcelReporter.Implementations.Providers.DataItemValueProviders;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Data;
using NSubstitute;

namespace ExcelReporter.Tests.Implementations.Providers.DataItemValueProviders
{
    [TestClass]
    public class DataItemValueProviderFactoryTest
    {
        [TestMethod]
        public void TestCreate()
        {
            var factory = new DefaultDataItemValueProviderFactory();
            Assert.IsInstanceOfType(factory.Create(null), typeof(ObjectPropertyValueProvider));

            var dataTable = new DataTable();
            dataTable.Columns.Add("Column", typeof(int));
            dataTable.Rows.Add(1);
            Assert.IsInstanceOfType(factory.Create(dataTable.Rows[0]), typeof(DataRowValueProvider));

            Assert.IsInstanceOfType(factory.Create(Substitute.For<IDataReader>()), typeof(DataReaderValueProvider));

            Assert.IsInstanceOfType(factory.Create(new int()), typeof(ObjectPropertyValueProvider));
            Assert.IsInstanceOfType(factory.Create(new object()), typeof(ObjectPropertyValueProvider));
        }
    }
}