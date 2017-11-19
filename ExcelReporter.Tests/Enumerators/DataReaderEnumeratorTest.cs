using ExcelReporter.Enumerators;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using ExcelReporter.Tests.CustomAsserts;

namespace ExcelReporter.Tests.Enumerators
{
    [TestClass]
    public class DataReaderEnumeratorTest
    {
        private readonly string _conStr = ConfigurationManager.ConnectionStrings["TestDb"].ConnectionString;

        public DataReaderEnumeratorTest()
        {
            TestHelper.InitDataDirectory();
        }

        [TestMethod]
        public void TestEnumerator()
        {
            ExceptionAssert.Throws<ArgumentNullException>(() => new DataReaderEnumerator(null));

            int counter = 0;
            IDataReader reader = Substitute.For<IDataReader>();
            reader.Read().Returns(x =>
            {
                counter++;
                return counter <= 3;
            });

            var enumerator = new DataReaderEnumerator(reader);
            ExceptionAssert.Throws<InvalidOperationException>(() => { _ = enumerator.Current; }, "Enumerator has not been started. Call MoveNext() method.");
            ExceptionAssert.Throws<InvalidOperationException>(() => { _ = enumerator.Current; }, "Enumerator has not been started. Call MoveNext() method.");
            while (enumerator.MoveNext())
            {
            }

            ExceptionAssert.Throws<InvalidOperationException>(() => enumerator.MoveNext(), "Enumerator has been finished");
            ExceptionAssert.Throws<InvalidOperationException>(() => enumerator.MoveNext(), "Enumerator has been finished");
            ExceptionAssert.Throws<InvalidOperationException>(() => { _ = enumerator.Current; }, "Enumerator has been finished");
            ExceptionAssert.Throws<InvalidOperationException>(() => { _ = enumerator.Current; }, "Enumerator has been finished");
            reader.Received(4).Read();

            ExceptionAssert.Throws<NotSupportedException>(() => enumerator.Reset(), $"{nameof(DataReaderEnumerator)} does not support reset method");

            reader.DidNotReceive().Close();
            enumerator.Dispose();
            reader.Received(1).Close();

            reader.IsClosed.Returns(true);

            ExceptionAssert.Throws<InvalidOperationException>(() => enumerator.MoveNext(), "DataReader has been closed");
            ExceptionAssert.Throws<InvalidOperationException>(() => { _ = enumerator.Current; }, "DataReader has been closed");
        }

        [TestMethod]
        public void TestEnumeratorWithRealSqlReader()
        {
            IDataReader reader = GetTestData();
            var enumerator = new DataReaderEnumerator(reader);

            Assert.IsTrue(enumerator.MoveNext());
            DataRow dataRow = enumerator.Current;
            Assert.AreEqual(1, dataRow.ItemArray[dataRow.Table.Columns["Id"].Ordinal]);
            Assert.AreEqual("Customer 1", dataRow.ItemArray[dataRow.Table.Columns["Name"].Ordinal]);
            Assert.AreEqual(false, dataRow.ItemArray[dataRow.Table.Columns["IsVip"].Ordinal]);
            Assert.AreEqual(DBNull.Value, dataRow.ItemArray[dataRow.Table.Columns["Type"].Ordinal]);

            Assert.IsTrue(enumerator.MoveNext());
            dataRow = enumerator.Current;
            Assert.AreEqual(2, dataRow.ItemArray[dataRow.Table.Columns["Id"].Ordinal]);
            Assert.AreEqual("Customer 2", dataRow.ItemArray[dataRow.Table.Columns["Name"].Ordinal]);
            Assert.AreEqual(true, dataRow.ItemArray[dataRow.Table.Columns["IsVip"].Ordinal]);
            Assert.AreEqual(1, dataRow.ItemArray[dataRow.Table.Columns["Type"].Ordinal]);

            dataRow = enumerator.Current;
            Assert.AreEqual(2, dataRow.ItemArray[dataRow.Table.Columns["Id"].Ordinal]);
            Assert.AreEqual("Customer 2", dataRow.ItemArray[dataRow.Table.Columns["Name"].Ordinal]);
            Assert.AreEqual(true, dataRow.ItemArray[dataRow.Table.Columns["IsVip"].Ordinal]);
            Assert.AreEqual(1, dataRow.ItemArray[dataRow.Table.Columns["Type"].Ordinal]);

            Assert.IsTrue(enumerator.MoveNext());
            dataRow = enumerator.Current;
            Assert.AreEqual(3, dataRow.ItemArray[dataRow.Table.Columns["Id"].Ordinal]);
            Assert.AreEqual("Customer 3", dataRow.ItemArray[dataRow.Table.Columns["Name"].Ordinal]);
            Assert.AreEqual(DBNull.Value, dataRow.ItemArray[dataRow.Table.Columns["IsVip"].Ordinal]);
            Assert.AreEqual(DBNull.Value, dataRow.ItemArray[dataRow.Table.Columns["Type"].Ordinal]);

            Assert.IsFalse(enumerator.MoveNext());
            ExceptionAssert.Throws<InvalidOperationException>(() => { _ = enumerator.Current; }, "Enumerator has been finished");
            ExceptionAssert.Throws<InvalidOperationException>(() => enumerator.MoveNext(), "Enumerator has been finished");

            enumerator.Dispose();
        }

        [TestMethod]
        public void TestEmptyEnumeratorWithRealSqlReader()
        {
            IDataReader reader = GetEmptyDataReader();
            var enumerator = new DataReaderEnumerator(reader);

            Assert.IsFalse(enumerator.MoveNext());
            ExceptionAssert.Throws<InvalidOperationException>(() => { _ = enumerator.Current; }, "Enumerator has been finished");
            ExceptionAssert.Throws<InvalidOperationException>(() => enumerator.MoveNext(), "Enumerator has been finished");

            enumerator.Dispose();
        }

        private IDataReader GetTestData()
        {
            IDbConnection connection = new SqlConnection(_conStr);
            IDbCommand command = connection.CreateCommand();
            command.CommandText = "SELECT Id, Name, IsVip, Type FROM Customers ORDER BY Id";
            connection.Open();
            return command.ExecuteReader(CommandBehavior.CloseConnection);
        }

        private IDataReader GetEmptyDataReader()
        {
            IDbConnection connection = new SqlConnection(_conStr);
            IDbCommand command = connection.CreateCommand();
            command.CommandText = "SELECT Id, Name, IsVip, Type FROM Customers WHERE 1 <> 1";
            connection.Open();
            return command.ExecuteReader(CommandBehavior.CloseConnection);
        }
    }
}