using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using ExcelReporter.Enumerators;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;

namespace ExcelReporter.Tests.Enumerators
{
    [TestClass]
    public class EnumeratorFactoryTest
    {
        [TestMethod]
        public void TestCreate()
        {
            Assert.IsNull(EnumeratorFactory.Create(null));

            IEnumerator enumerator = EnumeratorFactory.Create(new List<string>());
            Assert.IsTrue(enumerator.GetType().IsGenericType);
            Assert.AreEqual(1, enumerator.GetType().GetGenericArguments().Length);
            Assert.AreEqual("String", enumerator.GetType().GetGenericArguments().First().Name);

            enumerator = EnumeratorFactory.Create(new int[0]);
            Assert.AreEqual("SZArrayEnumerator", enumerator.GetType().Name);

            var dataReader = Substitute.For<IDataReader>();
            Assert.IsInstanceOfType(EnumeratorFactory.Create(dataReader), typeof(DataReaderEnumerator));

            enumerator = EnumeratorFactory.Create(new DataTable());
            Assert.IsTrue(enumerator.GetType().IsGenericType);
            Assert.AreEqual(1, enumerator.GetType().GetGenericArguments().Length);
            Assert.AreEqual("DataRow", enumerator.GetType().GetGenericArguments().First().Name);

            var dataSet = new DataSet();
            dataSet.Tables.Add(new DataTable());
            Assert.IsInstanceOfType(EnumeratorFactory.Create(dataSet), typeof(DataSetEnumerator));

            enumerator = EnumeratorFactory.Create(new object());
            Assert.AreEqual("SZArrayEnumerator", enumerator.GetType().Name);

            int counter = 0;
            while (enumerator.MoveNext())
            {
                counter++;
            }

            Assert.AreEqual(1, counter);
        }   
    }
}