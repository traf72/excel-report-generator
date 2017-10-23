using ExcelReporter.Iterators;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Data;

namespace ExcelReporter.Tests.Iterators
{
    [TestClass]
    public class DataSetIteratorTest
    {
        [TestMethod]
        public void TestIterator()
        {
            MyAssert.Throws<ArgumentNullException>(() => new DataSetIterator(null));

            DataSet dataSet = new DataSet();
            MyAssert.Throws<InvalidOperationException>(() => new DataSetIterator(dataSet), "DataSet does not contain any table");

            DataTable dataTable1 = new DataTable("Table1");
            dataTable1.Columns.Add("Column", typeof(int));

            dataTable1.Rows.Add(1);
            dataTable1.Rows.Add(2);
            dataTable1.Rows.Add(3);

            DataTable dataTable2 = new DataTable("Table2");
            dataTable2.Columns.Add("Column", typeof(int));

            dataTable2.Rows.Add(111);
            dataTable2.Rows.Add(222);
            dataTable2.Rows.Add(333);

            dataSet.Tables.Add(dataTable1);
            dataSet.Tables.Add(dataTable2);
            MyAssert.Throws<InvalidOperationException>(() => new DataSetIterator(dataSet, "BadTable"), "DataSet does not contain table with name \"BadTable\"");

            IList<int> result = new List<int>(dataTable1.Rows.Count);
            var iterator = new DataSetIterator(dataSet);
            MyAssert.Throws<InvalidOperationException>(() => iterator.Next(), "Iterator has not been started");
            MyAssert.Throws<InvalidOperationException>(() => iterator.Next(), "Iterator has not been started");
            while (iterator.HaxNext())
            {
                result.Add((int)iterator.Next().ItemArray[0]);
            }

            MyAssert.Throws<InvalidOperationException>(() => iterator.Next(), "Iterator has been finished");
            MyAssert.Throws<InvalidOperationException>(() => iterator.Next(), "Iterator has been finished");
            Assert.IsFalse(iterator.HaxNext());
            Assert.IsFalse(iterator.HaxNext());
            Assert.AreEqual(3, result.Count);
            Assert.AreEqual(1, result[0]);
            Assert.AreEqual(2, result[1]);
            Assert.AreEqual(3, result[2]);

            result.Clear();
            iterator = new DataSetIterator(dataSet, "Table2");
            MyAssert.Throws<InvalidOperationException>(() => iterator.Next(), "Iterator has not been started");
            MyAssert.Throws<InvalidOperationException>(() => iterator.Next(), "Iterator has not been started");
            while (iterator.HaxNext())
            {
                result.Add((int)iterator.Next().ItemArray[0]);
            }

            MyAssert.Throws<InvalidOperationException>(() => iterator.Next(), "Iterator has been finished");
            MyAssert.Throws<InvalidOperationException>(() => iterator.Next(), "Iterator has been finished");
            Assert.IsFalse(iterator.HaxNext());
            Assert.IsFalse(iterator.HaxNext());
            Assert.AreEqual(3, result.Count);
            Assert.AreEqual(111, result[0]);
            Assert.AreEqual(222, result[1]);
            Assert.AreEqual(333, result[2]);

            iterator.Reset();
            Assert.IsTrue(iterator.HaxNext());
            Assert.IsTrue(iterator.HaxNext());
            Assert.AreEqual(222, (int)iterator.Next().ItemArray[0]);

            result.Clear();
            dataTable1.Clear();
            iterator = new DataSetIterator(dataSet);
            MyAssert.Throws<InvalidOperationException>(() => iterator.Next(), "Iterator has not been started");
            MyAssert.Throws<InvalidOperationException>(() => iterator.Next(), "Iterator has not been started");
            while (iterator.HaxNext())
            {
                result.Add((int)iterator.Next().ItemArray[0]);
            }

            MyAssert.Throws<InvalidOperationException>(() => iterator.Next(), "Iterator has been finished");
            MyAssert.Throws<InvalidOperationException>(() => iterator.Next(), "Iterator has been finished");
            Assert.IsFalse(iterator.HaxNext());
            Assert.IsFalse(iterator.HaxNext());
            Assert.AreEqual(0, result.Count);
        }
    }
}