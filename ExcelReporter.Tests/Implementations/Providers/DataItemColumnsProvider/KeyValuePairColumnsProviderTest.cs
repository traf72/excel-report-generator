using ExcelReporter.Implementations.Providers.DataItemColumnsProvider;
using ExcelReporter.Interfaces.Providers.DataItemColumnsProvider;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Data;

namespace ExcelReporter.Tests.Implementations.Providers.DataItemColumnsProvider
{
    [TestClass]
    public class KeyValuePairColumnsProviderTest
    {
        [TestMethod]
        public void TestGetColumnsList()
        {
            IDataItemColumnsProvider columnsProvider = new KeyValuePairColumnsProvider();
            IList<ExcelDynamicColumn> columns = columnsProvider.GetColumnsList(new KeyValuePair<int, string>());

            Assert.AreEqual(2, columns.Count);

            Assert.AreEqual("Key", columns[0].Name);
            Assert.AreEqual("Key", columns[0].Caption);
            Assert.IsNull(columns[0].Width);

            Assert.AreEqual("Value", columns[1].Name);
            Assert.AreEqual("Value", columns[1].Caption);
            Assert.IsNull(columns[1].Width);

            columns = columnsProvider.GetColumnsList(new[] { new KeyValuePair<string, object>() });

            Assert.AreEqual(2, columns.Count);

            Assert.AreEqual("Key", columns[0].Name);
            Assert.AreEqual("Key", columns[0].Caption);
            Assert.IsNull(columns[0].Width);

            Assert.AreEqual("Value", columns[1].Name);
            Assert.AreEqual("Value", columns[1].Caption);
            Assert.IsNull(columns[1].Width);

            columns = columnsProvider.GetColumnsList(null);

            Assert.AreEqual(2, columns.Count);

            Assert.AreEqual("Key", columns[0].Name);
            Assert.AreEqual("Key", columns[0].Caption);
            Assert.IsNull(columns[0].Width);

            Assert.AreEqual("Value", columns[1].Name);
            Assert.AreEqual("Value", columns[1].Caption);
            Assert.IsNull(columns[1].Width);

            columns = columnsProvider.GetColumnsList(new DataSet());

            Assert.AreEqual(2, columns.Count);

            Assert.AreEqual("Key", columns[0].Name);
            Assert.AreEqual("Key", columns[0].Caption);
            Assert.IsNull(columns[0].Width);

            Assert.AreEqual("Value", columns[1].Name);
            Assert.AreEqual("Value", columns[1].Caption);
            Assert.IsNull(columns[1].Width);
        }
    }
}