using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using ExcelReporter.Exceptions;
using ExcelReporter.Rendering.Providers.DataItemValueProviders;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;
using NSubstitute.ExceptionExtensions;

namespace ExcelReporter.Tests.Rendering.Providers.DataItemValueProviders
{
    [TestClass]
    public class DataReaderValueProviderTest
    {
        private readonly string _conStr = ConfigurationManager.ConnectionStrings["TestDb"].ConnectionString;

        public DataReaderValueProviderTest()
        {
            TestHelper.InitDataDirectory();
        }

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

            ExceptionAssert.Throws<ColumnNotFoundException>(() => provider.GetValue("BadColumn", dataReader), "DataReader does not contain column \"BadColumn\"");
            ExceptionAssert.Throws<ColumnNotFoundException>(() => provider.GetValue("BadColumn2", dataReader), "DataReader does not contain column \"BadColumn2\"");
            ExceptionAssert.Throws<ArgumentException>(() => provider.GetValue(null, dataReader));
            ExceptionAssert.Throws<ArgumentException>(() => provider.GetValue(string.Empty, dataReader));
            ExceptionAssert.Throws<ArgumentException>(() => provider.GetValue(" ", dataReader));
            ExceptionAssert.Throws<ArgumentNullException>(() => provider.GetValue("Column1", null));

            dataReader.IsClosed.Returns(true);
            ExceptionAssert.Throws<InvalidOperationException>(() => provider.GetValue("Column1", dataReader), "DataReader is closed");
        }

        [TestMethod]
        public void TestGetValueWithRealSqlReader()
        {
            IGenericDataItemValueProvider<IDataReader> provider = new DataReaderValueProvider();
            IDataReader reader = GetTestData();

            reader.Read();
            Assert.AreEqual(1, provider.GetValue("Id", reader));
            Assert.AreEqual(1, provider.GetValue("id", reader));
            Assert.AreEqual("Customer 1", provider.GetValue("Name", reader));
            Assert.AreEqual(false, provider.GetValue("IsVip", reader));
            Assert.IsNull(provider.GetValue("Type", reader));

            reader.Read();
            Assert.AreEqual(2, provider.GetValue("Id", reader));
            Assert.AreEqual("Customer 2", provider.GetValue("Name", reader));
            Assert.AreEqual(true, provider.GetValue("IsVip", reader));
            Assert.AreEqual(1, provider.GetValue("Type", reader));

            reader.Read();
            Assert.AreEqual(3, provider.GetValue("Id", reader));
            Assert.AreEqual("Customer 3", provider.GetValue("Name", reader));
            Assert.IsNull(provider.GetValue("IsVip", reader));
            Assert.IsNull(provider.GetValue("Type", reader));

            reader.Close();
        }

        private IDataReader GetTestData()
        {
            IDbConnection connection = new SqlConnection(_conStr);
            IDbCommand command = connection.CreateCommand();
            command.CommandText = "SELECT Id, Name, IsVip, Type FROM Customers ORDER BY Id";
            connection.Open();
            return command.ExecuteReader(CommandBehavior.CloseConnection);
        }
    }
}