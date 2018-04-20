using ExcelReportGenerator.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Data;
using System.Data.SqlClient;

namespace ExcelReportGenerator.Tests.Extensions
{
    [TestClass]
    public class DataReaderExtensionsTest
    {
        private readonly string _conStr = Configuration.TestDbConectionString;

        [TestMethod]
        public void TestSafeGetValue()
        {
            IDataReader reader = GetTestBoolData();
            reader.Read();

            Assert.IsNull(reader.SafeGetValue(reader.GetOrdinal("IsVip")));
            object id = reader.SafeGetValue(reader.GetOrdinal("Id"));
            Assert.IsNotNull(id);
            Assert.AreNotEqual(0, id);

            reader.Close();

            reader = GetTestIntData();
            reader.Read();

            Assert.IsNull(reader.SafeGetValue(reader.GetOrdinal("Type")));
            id = reader.SafeGetValue(reader.GetOrdinal("Id"));
            Assert.IsNotNull(id);
            Assert.AreNotEqual(0, id);

            reader.Close();

            reader = GetTestStringData();
            reader.Read();

            Assert.IsNull(reader.SafeGetValue(reader.GetOrdinal("Description")));
            id = reader.SafeGetValue(reader.GetOrdinal("Id"));
            Assert.IsNotNull(id);
            Assert.AreNotEqual(0, id);

            reader.Close();
        }

        private IDataReader GetTestBoolData()
        {
            IDbConnection connection = new SqlConnection(_conStr);
            IDbCommand command = connection.CreateCommand();
            command.CommandText = "SELECT TOP 1 * FROM Customers WHERE IsVip IS NULL";
            connection.Open();
            return command.ExecuteReader(CommandBehavior.CloseConnection);
        }

        private IDataReader GetTestIntData()
        {
            IDbConnection connection = new SqlConnection(_conStr);
            IDbCommand command = connection.CreateCommand();
            command.CommandText = "SELECT TOP 1 * FROM Customers WHERE Type IS NULL";
            connection.Open();
            return command.ExecuteReader(CommandBehavior.CloseConnection);
        }

        private IDataReader GetTestStringData()
        {
            IDbConnection connection = new SqlConnection(_conStr);
            IDbCommand command = connection.CreateCommand();
            command.CommandText = "SELECT TOP 1 * FROM Customers WHERE Description IS NULL";
            connection.Open();
            return command.ExecuteReader(CommandBehavior.CloseConnection);
        }
    }
}