using ExcelReportGenerator.Enumerators;
using ExcelReportGenerator.Tests.Rendering;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Data;

namespace ExcelReportGenerator.Tests.Enumerators
{
    public class DataReaderEnumeratorTest
    {
        [Test]
        public void TestEnumerator()
        {
            IDataReader reader = new DataProvider().GetAllCustomersDataReader();
            var enumerator = new DataReaderEnumerator(reader);

            Assert.IsTrue(reader.IsClosed);

            Assert.AreEqual(3, enumerator.RowCount);

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

            Assert.AreEqual(3, enumerator.RowCount);

            enumerator.Reset();

            Assert.AreEqual(3, enumerator.RowCount);

            IList<DataRow> result = new List<DataRow>();
            while (enumerator.MoveNext())
            {
                result.Add(enumerator.Current);
            }

            Assert.AreEqual(3, result.Count);
            Assert.AreEqual(3, enumerator.RowCount);
        }

        [Test]
        public void TestEmptyEnumeratorWithRealSqlReader()
        {
            IDataReader reader = new DataProvider().GetEmptyDataReader();
            var enumerator = new DataReaderEnumerator(reader);

            Assert.IsTrue(reader.IsClosed);
            Assert.AreEqual(0, enumerator.RowCount);
            Assert.IsFalse(enumerator.MoveNext());
            Assert.AreEqual(0, enumerator.RowCount);
        }
    }
}