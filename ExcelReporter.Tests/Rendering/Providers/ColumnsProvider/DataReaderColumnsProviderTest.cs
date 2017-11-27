using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using ExcelReporter.Rendering;
using ExcelReporter.Rendering.Providers.ColumnsProviders;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Rendering.Providers.ColumnsProvider
{
    [TestClass]
    public class DataReaderColumnsProviderTest
    {
        private readonly string _conStr = ConfigurationManager.ConnectionStrings["TestDb"].ConnectionString;

        public DataReaderColumnsProviderTest()
        {
            TestHelper.InitDataDirectory();
        }

        [TestMethod]
        public void TestGetColumnsList()
        {
            IDataReader dataReader = GetTestData();
            IDataItemColumnsProvider columnsProvider = new DataReaderColumnsProvider();
            IList<ExcelDynamicColumn> columns = columnsProvider.GetColumnsList(dataReader);

            Assert.AreEqual(4, columns.Count);

            Assert.AreEqual("Id", columns[0].Name);
            Assert.AreEqual("Id", columns[0].Caption);
            Assert.IsNull(columns[0].Width);

            Assert.AreEqual("Name", columns[1].Name);
            Assert.AreEqual("Name", columns[1].Caption);
            Assert.IsNull(columns[1].Width);

            Assert.AreEqual("IsVip", columns[2].Name);
            Assert.AreEqual("IsVip", columns[2].Caption);
            Assert.IsNull(columns[2].Width);

            Assert.AreEqual("Type", columns[3].Name);
            Assert.AreEqual("Type", columns[3].Caption);
            Assert.IsNull(columns[3].Width);

            dataReader.Close();
        }

        [TestMethod]
        public void TestGetColumnsListIfDataReaderIsEmpty()
        {
            IDataReader dataReader = GetEmptyDataReader();
            IDataItemColumnsProvider columnsProvider = new DataReaderColumnsProvider();
            IList<ExcelDynamicColumn> columns = columnsProvider.GetColumnsList(dataReader);

            Assert.AreEqual(4, columns.Count);

            Assert.AreEqual("Id", columns[0].Name);
            Assert.AreEqual("Id", columns[0].Caption);
            Assert.IsNull(columns[0].Width);

            Assert.AreEqual("Name", columns[1].Name);
            Assert.AreEqual("Name", columns[1].Caption);
            Assert.IsNull(columns[1].Width);

            Assert.AreEqual("IsVip", columns[2].Name);
            Assert.AreEqual("IsVip", columns[2].Caption);
            Assert.IsNull(columns[2].Width);

            Assert.AreEqual("Type", columns[3].Name);
            Assert.AreEqual("Type", columns[3].Caption);
            Assert.IsNull(columns[3].Width);

            dataReader.Close();
        }

        [TestMethod]
        public void TestGetColumnsListIfDataReaderIsNull()
        {
            IDataItemColumnsProvider columnsProvider = new DataReaderColumnsProvider();
            Assert.AreEqual(0, columnsProvider.GetColumnsList(null).Count);
        }

        private IDataReader GetTestData()
        {
            IDbConnection connection = new SqlConnection(_conStr);
            IDbCommand command = connection.CreateCommand();
            command.CommandText = "SELECT Id, Name, IsVip, Type FROM Customers";
            connection.Open();
            return command.ExecuteReader(CommandBehavior.CloseConnection);
        }

        private IDataReader GetEmptyDataReader()
        {
            IDbConnection connection = new SqlConnection(_conStr);
            IDbCommand command = connection.CreateCommand();
            command.CommandText = "SELECT Id, Name, IsVip, Type FROM Customers WHERE 1 <> 1";
            connection.Open();
            return command.ExecuteReader(CommandBehavior.CloseConnection);
        }
    }
}