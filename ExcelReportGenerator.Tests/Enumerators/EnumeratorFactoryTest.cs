using ExcelReportGenerator.Enumerators;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;
using System.Collections;
using System.Collections.Generic;
using System.Data;

namespace ExcelReportGenerator.Tests.Enumerators
{
    [TestClass]
    public class EnumeratorFactoryTest
    {
        [TestMethod]
        public void TestCreate()
        {
            Assert.IsNull(EnumeratorFactory.Create(null));

            Assert.IsInstanceOfType(EnumeratorFactory.Create(new List<string>()), typeof(EnumerableEnumerator));
            Assert.IsInstanceOfType(EnumeratorFactory.Create(new int[0]), typeof(EnumerableEnumerator));
            Assert.IsInstanceOfType(EnumeratorFactory.Create(new Dictionary<string, object>()), typeof(EnumerableEnumerator));
            Assert.IsInstanceOfType(EnumeratorFactory.Create(new HashSet<string>()), typeof(EnumerableEnumerator));
            Assert.IsInstanceOfType(EnumeratorFactory.Create(new Hashtable()), typeof(EnumerableEnumerator));
            Assert.IsInstanceOfType(EnumeratorFactory.Create(new ArrayList()), typeof(EnumerableEnumerator));

            var dataSet = new DataSet();
            dataSet.Tables.Add(new DataTable());

            Assert.IsInstanceOfType(EnumeratorFactory.Create(dataSet), typeof(DataSetEnumerator));
            Assert.IsInstanceOfType(EnumeratorFactory.Create(new DataTable()), typeof(DataTableEnumerator));

            var dataReader = Substitute.For<IDataReader>();
            Assert.IsInstanceOfType(EnumeratorFactory.Create(dataReader), typeof(DataReaderEnumerator));
        }
    }
}