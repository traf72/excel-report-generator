using System;
using System.Data;
using ExcelReportGenerator.Exceptions;
using ExcelReportGenerator.Rendering.Providers.DataItemValueProviders;
using ExcelReportGenerator.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReportGenerator.Tests.Rendering.Providers.DataItemValueProviders
{
    [TestClass]
    public class DataRowValueProviderTest
    {
        [TestMethod]
        public void TestGetValue()
        {
            IGenericDataItemValueProvider<DataRow> provider = new DataRowValueProvider();

            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("Column1", typeof(int));
            dataTable.Columns.Add("Column2", typeof(string));

            dataTable.Rows.Add(1, "One");
            dataTable.Rows.Add(2, "Two");
            dataTable.Rows.Add(3, "Three");

            Assert.AreEqual(dataTable.Rows[0].ItemArray[0], provider.GetValue("Column1", dataTable.Rows[0]));
            Assert.AreEqual(dataTable.Rows[0].ItemArray[0], provider.GetValue(" Column1 ", dataTable.Rows[0]));
            Assert.AreEqual(dataTable.Rows[0].ItemArray[0], provider.GetValue(" column1 ", dataTable.Rows[0]));
            Assert.AreEqual(dataTable.Rows[0].ItemArray[1], provider.GetValue("Column2", dataTable.Rows[0]));
            Assert.AreEqual(dataTable.Rows[1].ItemArray[0], provider.GetValue("Column1", dataTable.Rows[1]));
            Assert.AreEqual(dataTable.Rows[1].ItemArray[1], provider.GetValue("Column2", dataTable.Rows[1]));
            Assert.AreEqual(dataTable.Rows[2].ItemArray[0], provider.GetValue("Column1", dataTable.Rows[2]));
            Assert.AreEqual(dataTable.Rows[2].ItemArray[1], provider.GetValue("Column2", dataTable.Rows[2]));

            ExceptionAssert.Throws<ColumnNotFoundException>(() => provider.GetValue("BadColumn", dataTable.Rows[0]), "DataRow does not contain column \"BadColumn\"");
            ExceptionAssert.Throws<ArgumentException>(() => provider.GetValue(null, dataTable.Rows[0]));
            ExceptionAssert.Throws<ArgumentException>(() => provider.GetValue(string.Empty, dataTable.Rows[0]));
            ExceptionAssert.Throws<ArgumentException>(() => provider.GetValue(" ", dataTable.Rows[0]));
            ExceptionAssert.Throws<ArgumentNullException>(() => provider.GetValue("Column1", null));
        }
    }
}