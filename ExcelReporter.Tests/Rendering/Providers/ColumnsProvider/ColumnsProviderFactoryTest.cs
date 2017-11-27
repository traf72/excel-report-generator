using System.Collections;
using System.Collections.Generic;
using System.Data;
using ExcelReporter.Rendering.Providers.ColumnsProviders;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;

namespace ExcelReporter.Tests.Rendering.Providers.ColumnsProvider
{
    [TestClass]
    public class ColumnsProviderFactoryTest
    {
        [TestMethod]
        public void TestCreate()
        {
            IDataItemColumnsProviderFactory factory = new ColumnsProviderFactory();

            Assert.IsNull(factory.Create(null));
            Assert.AreEqual(typeof(DataReaderColumnsProvider), factory.Create(Substitute.For<IDataReader>()).GetType());
            Assert.AreEqual(typeof(DataTableColumnsProvider), factory.Create(new DataTable()).GetType());
            Assert.AreEqual(typeof(DataSetColumnsProvider), factory.Create(new DataSet()).GetType());

            Assert.AreEqual(typeof(KeyValuePairColumnsProvider), factory.Create(new KeyValuePair<string, int>()).GetType());
            Assert.AreEqual(typeof(KeyValuePairColumnsProvider), factory.Create(new KeyValuePair<int, double>()).GetType());
            Assert.AreEqual(typeof(KeyValuePairColumnsProvider), factory.Create(new[] { new KeyValuePair<string, int>() }).GetType());
            Assert.AreEqual(typeof(KeyValuePairColumnsProvider), factory.Create(new List<KeyValuePair<decimal, string>> { new KeyValuePair<decimal, string>() }).GetType());
            Assert.AreEqual(typeof(KeyValuePairColumnsProvider), factory.Create(new Dictionary<string, TypeColumnsProviderTest.TestType>()).GetType());

            Assert.AreEqual(typeof(DictionaryColumnsProvider), factory.Create(new[] { new Dictionary<string, int>() }).GetType());
            Assert.AreEqual(typeof(DictionaryColumnsProvider), factory.Create(new List<IDictionary<string, string>> { new Dictionary<string, string>() }).GetType());
            Assert.AreEqual(typeof(DictionaryColumnsProvider), factory.Create(new List<IDictionary<string, decimal>> { new Dictionary<string, decimal>() }).GetType());

            Assert.AreEqual(typeof(GenericEnumerableColumnsProvider), factory.Create(new List<IDictionary<object, decimal>> { new Dictionary<object, decimal>() }).GetType());
            Assert.AreEqual(typeof(GenericEnumerableColumnsProvider), factory.Create(new List<IDictionary<int, decimal>> { new Dictionary<int, decimal>() }).GetType());
            Assert.AreEqual(typeof(GenericEnumerableColumnsProvider), factory.Create(new[] { "str" }).GetType());
            Assert.AreEqual(typeof(GenericEnumerableColumnsProvider), factory.Create(new[] { 1, 2 }).GetType());
            Assert.AreEqual(typeof(GenericEnumerableColumnsProvider), factory.Create(new List<decimal> { 1m, 2m }).GetType());
            Assert.AreEqual(typeof(EnumerableColumnsProvider), factory.Create(new ArrayList { "str" }).GetType());
            Assert.AreEqual(typeof(GenericEnumerableColumnsProvider), factory.Create("str").GetType());

            Assert.AreEqual(typeof(ObjectColumnsProvider), factory.Create(1).GetType());
            Assert.AreEqual(typeof(ObjectColumnsProvider), factory.Create(new TypeColumnsProviderTest.TestType()).GetType());
        }
    }
}