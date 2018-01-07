using System.Collections.Generic;
using ExcelReportGenerator.Rendering;
using ExcelReportGenerator.Rendering.Providers.ColumnsProviders;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReportGenerator.Tests.Rendering.Providers.ColumnsProvider
{
    [TestClass]
    public class DictionaryColumnsProviderTest
    {
        [TestMethod]
        public void TestGetColumnsList()
        {
            IColumnsProvider columnsProvider = new DictionaryColumnsProvider<object>();

            IDictionary<string, object>[] dictArray = {
                new Dictionary<string, object> { ["Id"] = 1, ["Name"] = "One", ["IsVip"] = true },
                new Dictionary<string, object> { ["Id"] = 2, ["Name"] = "Two" },
            };

            IList<ExcelDynamicColumn> columns = columnsProvider.GetColumnsList(dictArray);

            Assert.AreEqual(3, columns.Count);

            Assert.AreEqual("Id", columns[0].Name);
            Assert.AreEqual("Id", columns[0].Caption);
            Assert.AreEqual(typeof(int), columns[0].DataType);
            Assert.IsNull(columns[0].Width);

            Assert.AreEqual("Name", columns[1].Name);
            Assert.AreEqual("Name", columns[1].Caption);
            Assert.AreEqual(typeof(string), columns[1].DataType);
            Assert.IsNull(columns[1].Width);

            Assert.AreEqual("IsVip", columns[2].Name);
            Assert.AreEqual("IsVip", columns[2].Caption);
            Assert.AreEqual(typeof(bool), columns[2].DataType);
            Assert.IsNull(columns[2].Width);

            columnsProvider = new DictionaryColumnsProvider<int>();

            IDictionary<string, int>[] dictArray2 = {
                new Dictionary<string, int> { ["Id"] = 2, ["Number"] = 3 },
                new Dictionary<string, int> { ["Id"] = 1, ["Number"] = 4, ["Number2"] = 5 },
            };

            columns = columnsProvider.GetColumnsList(dictArray2);

            Assert.AreEqual(2, columns.Count);

            Assert.AreEqual("Id", columns[0].Name);
            Assert.AreEqual("Id", columns[0].Caption);
            Assert.AreEqual(typeof(int), columns[0].DataType);
            Assert.IsNull(columns[0].Width);

            Assert.AreEqual("Number", columns[1].Name);
            Assert.AreEqual("Number", columns[1].Caption);
            Assert.AreEqual(typeof(int), columns[1].DataType);
            Assert.IsNull(columns[1].Width);
        }

        [TestMethod]
        public void TestGetColumnsListIfDictionaryIsNullOrEmpty()
        {
            IColumnsProvider columnsProvider = new DictionaryColumnsProvider<object>();
            Assert.AreEqual(0, columnsProvider.GetColumnsList(null).Count);
            Assert.AreEqual(0, columnsProvider.GetColumnsList(new Dictionary<string, object>[0]).Count);
            Assert.AreEqual(0, columnsProvider.GetColumnsList(new[] { new Dictionary<string, object>() }).Count);
        }
    }
}