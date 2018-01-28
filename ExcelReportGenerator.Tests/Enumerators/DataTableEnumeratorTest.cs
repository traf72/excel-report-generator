using ExcelReportGenerator.Enumerators;
using ExcelReportGenerator.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Data;

namespace ExcelReportGenerator.Tests.Enumerators
{
    [TestClass]
    public class DataTableEnumeratorTest
    {
        [TestMethod]
        public void TestEnumerator()
        {
            ExceptionAssert.Throws<ArgumentNullException>(() => new DataTableEnumerator(null));

            DataTable dataTable = new DataTable("Table1");
            dataTable.Columns.Add("Column", typeof(int));

            dataTable.Rows.Add(1);
            dataTable.Rows.Add(2);
            dataTable.Rows.Add(3);

            IList<int> result = new List<int>(dataTable.Rows.Count);
            var enumerator = new DataTableEnumerator(dataTable);

            Assert.AreEqual(3, enumerator.RowCount);

            Assert.IsNull(enumerator.Current);
            Assert.IsNull(enumerator.Current);
            while (enumerator.MoveNext())
            {
                result.Add((int)enumerator.Current.ItemArray[0]);
            }

            Assert.IsFalse(enumerator.MoveNext());
            Assert.IsFalse(enumerator.MoveNext());

            Assert.AreEqual(3, enumerator.Current.ItemArray[0]);
            Assert.AreEqual(3, enumerator.Current.ItemArray[0]);

            Assert.AreEqual(3, result.Count);
            Assert.AreEqual(1, result[0]);
            Assert.AreEqual(2, result[1]);
            Assert.AreEqual(3, result[2]);

            Assert.AreEqual(3, enumerator.RowCount);

            enumerator.Dispose();
            enumerator.Reset();
            result.Clear();

            Assert.AreEqual(3, enumerator.RowCount);

            Assert.IsNull(enumerator.Current);
            Assert.IsNull(enumerator.Current);
            while (enumerator.MoveNext())
            {
                result.Add((int)enumerator.Current.ItemArray[0]);
            }

            Assert.IsFalse(enumerator.MoveNext());
            Assert.IsFalse(enumerator.MoveNext());

            Assert.AreEqual(3, enumerator.Current.ItemArray[0]);
            Assert.AreEqual(3, enumerator.Current.ItemArray[0]);

            Assert.AreEqual(3, result.Count);
            Assert.AreEqual(1, result[0]);
            Assert.AreEqual(2, result[1]);
            Assert.AreEqual(3, result[2]);

            Assert.AreEqual(3, enumerator.RowCount);
        }
    }
}