using System;
using System.Data;
using ExcelReporter.Exceptions;
using ExcelReporter.Implementations.Providers.DataItemValueProviders;
using ExcelReporter.Interfaces.Providers;
using ExcelReporter.Interfaces.Providers.DataItemValueProviders;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;
using NSubstitute.ExceptionExtensions;

namespace ExcelReporter.Tests.Implementations.Providers.DataItemValueProviders
{
    [TestClass]
    public class DataReaderValueProviderTest
    {
        [TestMethod]
        public void TestGetValue()
        {
            IGenericDataItemValueProvider<IDataReader> provider = new DataReaderValueProvider();

            IDataReader dataReader = Substitute.For<IDataReader>();
            dataReader.GetOrdinal("Column1").Returns(0);
            dataReader.GetOrdinal("column1").Returns(0);
            dataReader.GetOrdinal("Column2").Returns(1);
            dataReader.GetOrdinal("BadColumn").Returns(-1);
            dataReader.GetOrdinal("BadColumn2").Throws(new IndexOutOfRangeException());

            dataReader.GetValue(0).Returns(5);
            dataReader.GetValue(1).Returns("Five");

            Assert.AreEqual(5, provider.GetValue("Column1", dataReader));
            Assert.AreEqual(5, provider.GetValue("column1", dataReader));
            Assert.AreEqual(5, provider.GetValue(" column1 ", dataReader));
            Assert.AreEqual("Five", provider.GetValue("Column2", dataReader));

            MyAssert.Throws<ColumnNotFoundException>(() => provider.GetValue("BadColumn", dataReader), "DataReader does not contain column \"BadColumn\"");
            MyAssert.Throws<ColumnNotFoundException>(() => provider.GetValue("BadColumn2", dataReader), "DataReader does not contain column \"BadColumn2\"");
            MyAssert.Throws<ArgumentException>(() => provider.GetValue(null, dataReader));
            MyAssert.Throws<ArgumentException>(() => provider.GetValue(string.Empty, dataReader));
            MyAssert.Throws<ArgumentException>(() => provider.GetValue(" ", dataReader));
            MyAssert.Throws<ArgumentNullException>(() => provider.GetValue("Column1", null));

            dataReader.IsClosed.Returns(true);
            MyAssert.Throws<InvalidOperationException>(() => provider.GetValue("Column1", dataReader), "DataReader is closed");
        }
    }
}