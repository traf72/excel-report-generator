using ExcelReportGenerator.Enumerators;
using ExcelReportGenerator.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Data;

namespace ExcelReportGenerator.Tests.Enumerators
{
    [TestClass]
    public class DataSetEnumeratorTest
    {
        [TestMethod]
        public void TestEnumerator()
        {
            ExceptionAssert.Throws<ArgumentNullException>(() => new DataSetEnumerator(null));

            DataSet dataSet = new DataSet();
            ExceptionAssert.Throws<InvalidOperationException>(() => new DataSetEnumerator(dataSet), "DataSet does not contain any table");

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
            ExceptionAssert.Throws<InvalidOperationException>(() => new DataSetEnumerator(dataSet, "BadTable"), "DataSet does not contain table with name \"BadTable\"");

            IList<int> result = new List<int>(dataTable1.Rows.Count);
            var enumerator = new DataSetEnumerator(dataSet);

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

            enumerator.Reset();
            result.Clear();

            enumerator = new DataSetEnumerator(dataSet, "Table2");
            Assert.IsNull(enumerator.Current);
            Assert.IsNull(enumerator.Current);
            while (enumerator.MoveNext())
            {
                result.Add((int)enumerator.Current.ItemArray[0]);
            }

            Assert.IsFalse(enumerator.MoveNext());
            Assert.IsFalse(enumerator.MoveNext());

            Assert.AreEqual(333, enumerator.Current.ItemArray[0]);
            Assert.AreEqual(333, enumerator.Current.ItemArray[0]);
            Assert.AreEqual(3, result.Count);
            Assert.AreEqual(111, result[0]);
            Assert.AreEqual(222, result[1]);
            Assert.AreEqual(333, result[2]);
        }
    }
}