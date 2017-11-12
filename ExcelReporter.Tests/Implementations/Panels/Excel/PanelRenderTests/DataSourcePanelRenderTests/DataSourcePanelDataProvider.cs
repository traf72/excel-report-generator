using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace ExcelReporter.Tests.Implementations.Panels.Excel.PanelRenderTests.DataSourcePanelRenderTests
{
    public class DataSourcePanelDataProvider
    {
        private class TestDataProvider
        {
            public TestItem GetSingleItem()
            {
                return new TestItem("Test", new DateTime(2017, 11, 1), 55.76m, new Contacts("15", "345"));
            }

            public TestItem GetNullItem()
            {
                return null;
            }

            public IEnumerable<TestItem> GetIEnumerable()
            {
                return new[]
                {
                    new TestItem("Test1", new DateTime(2017, 11, 1), 55.76m, new Contacts("15", "345")),
                    new TestItem("Test2", new DateTime(2017, 11, 2), 110m, new Contacts("76", "753465")),
                    new TestItem("Test3", new DateTime(2017, 11, 3), 5500.80m, new Contacts("1533", "5456")),
                };
            }

            public IEnumerable<TestItem> GetEmptyIEnumerable()
            {
                return Enumerable.Empty<TestItem>();
            }

            public IDataReader GetAllCustomersDataReader()
            {
                string conStr = ConfigurationManager.ConnectionStrings["TestDb"].ConnectionString;
                IDbConnection connection = new SqlConnection(conStr);
                IDbCommand command = connection.CreateCommand();
                command.CommandText = "SELECT * FROM Customers";
                connection.Open();
                return command.ExecuteReader(CommandBehavior.CloseConnection);
            }

            public IDataReader GetEmptyDataReader()
            {
                string conStr = ConfigurationManager.ConnectionStrings["TestDb"].ConnectionString;
                IDbConnection connection = new SqlConnection(conStr);
                IDbCommand command = connection.CreateCommand();
                command.CommandText = "SELECT * FROM Customers WHERE 1 <> 1";
                connection.Open();
                return command.ExecuteReader(CommandBehavior.CloseConnection);
            }
        }

        private class TestItem
        {
            public TestItem(string name, DateTime date, decimal sum, Contacts contacts = null)
            {
                Name = name;
                Date = date;
                Sum = sum;
                Contacts = contacts;
            }

            public string Name { get; set; }

            public DateTime Date { get; set; }

            public decimal Sum { get; set; }

            public Contacts Contacts { get; set; }
        }

        private class Contacts
        {
            public Contacts(string phone, string fax)
            {
                Phone = phone;
                Fax = fax;
            }

            public string Phone { get; set; }

            public string Fax { get; set; }

            public override string ToString()
            {
                return $"{Phone}_{Fax}";
            }
        }
    }
}