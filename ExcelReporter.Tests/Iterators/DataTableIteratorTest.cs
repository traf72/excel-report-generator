using ExcelReporter.Iterators;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Data;

namespace ExcelReporter.Tests.Iterators
{
    [TestClass]
    public class DataTableIteratorTest
    {
        [TestMethod]
        public void TestIterator()
        {
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("Column", typeof(int));

            dataTable.Rows.Add(1);
            dataTable.Rows.Add(2);
            dataTable.Rows.Add(3);

            IList<int> result = new List<int>(dataTable.Rows.Count);
            var iterator = new DataTableIterator(dataTable);
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

            iterator.Reset();
            Assert.IsTrue(iterator.HaxNext());
            Assert.IsTrue(iterator.HaxNext());
            Assert.AreEqual(2, (int)iterator.Next().ItemArray[0]);

            result.Clear();
            dataTable.Clear();
            iterator = new DataTableIterator(dataTable);
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

            MyAssert.Throws<ArgumentNullException>(() => new DataTableIterator(null));
        }
    }
}