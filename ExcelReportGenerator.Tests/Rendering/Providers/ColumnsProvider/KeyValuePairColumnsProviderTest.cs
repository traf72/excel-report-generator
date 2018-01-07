using ExcelReportGenerator.Rendering;
using ExcelReportGenerator.Rendering.Providers.ColumnsProviders;
using ExcelReportGenerator.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Data;

namespace ExcelReportGenerator.Tests.Rendering.Providers.ColumnsProvider
{
    [TestClass]
    public class KeyValuePairColumnsProviderTest
    {
        [TestMethod]
        public void TestGetColumnsList()
        {
            IColumnsProvider columnsProvider = new KeyValuePairColumnsProvider();
            IList<ExcelDynamicColumn> columns = columnsProvider.GetColumnsList(new KeyValuePair<int, string>());

            Assert.AreEqual(2, columns.Count);

            Assert.AreEqual("Key", columns[0].Name);
            Assert.AreEqual("Key", columns[0].Caption);
            Assert.AreEqual(typeof(int), columns[0].DataType);
            Assert.IsNull(columns[0].Width);

            Assert.AreEqual("Value", columns[1].Name);
            Assert.AreEqual("Value", columns[1].Caption);
            Assert.AreEqual(typeof(string), columns[1].DataType);
            Assert.IsNull(columns[1].Width);

            columns = columnsProvider.GetColumnsList(new[] { new KeyValuePair<Guid?, decimal>() });

            Assert.AreEqual(2, columns.Count);

            Assert.AreEqual("Key", columns[0].Name);
            Assert.AreEqual("Key", columns[0].Caption);
            Assert.AreEqual(typeof(Guid?), columns[0].DataType);
            Assert.IsNull(columns[0].Width);

            Assert.AreEqual("Value", columns[1].Name);
            Assert.AreEqual("Value", columns[1].Caption);
            Assert.AreEqual(typeof(decimal), columns[1].DataType);
            Assert.IsNull(columns[1].Width);

            columns = columnsProvider.GetColumnsList(null);

            Assert.AreEqual(2, columns.Count);

            Assert.AreEqual("Key", columns[0].Name);
            Assert.AreEqual("Key", columns[0].Caption);
            Assert.IsNull(columns[0].DataType);
            Assert.IsNull(columns[0].Width);

            Assert.AreEqual("Value", columns[1].Name);
            Assert.AreEqual("Value", columns[1].Caption);
            Assert.IsNull(columns[1].DataType);
            Assert.IsNull(columns[1].Width);

            ExceptionAssert.Throws<InvalidOperationException>(() => columnsProvider.GetColumnsList(new DataSet()), "Type of data must be KeyValuePair<TKey, TValue> or IEnumerable<KeyValuePair<TKey, TValue>>");
        }
    }
}