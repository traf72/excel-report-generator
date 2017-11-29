using ExcelReporter.Enumerators;
using ExcelReporter.Tests.CustomAsserts;
using ExcelReporter.Tests.Rendering;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;
using System;
using System.Data;

namespace ExcelReporter.Tests.Enumerators
{
    [TestClass]
    public class DataReaderEnumeratorTest
    {
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
            IDataReader reader = new DataProvider().GetAllCustomersDataReader();
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
            IDataReader reader = new DataProvider().GetEmptyDataReader();
            var enumerator = new DataReaderEnumerator(reader);

            Assert.IsFalse(enumerator.MoveNext());
            ExceptionAssert.Throws<InvalidOperationException>(() => { _ = enumerator.Current; }, "Enumerator has been finished");
            ExceptionAssert.Throws<InvalidOperationException>(() => enumerator.MoveNext(), "Enumerator has been finished");

            enumerator.Dispose();
        }
    }
}