﻿using System.Data;
using System.Data.SqlClient;
using ExcelReportGenerator.Attributes;
using ExcelReportGenerator.Enums;

namespace ExcelReportGenerator.Tests.Rendering;

public class DataProvider
{
    private static readonly string _conStr;

    private readonly IEnumerable<TestItem> _testData = new[]
    {
        new TestItem("Test1", new DateTime(2017, 11, 1), 55.76m, new Contacts("15", "345"))
        {
            Children = new List<ChildItem>
            {
                new("Test1_Child1_F1", "Test1_Child1_F2")
                {
                    Children = new[]
                    {
                        new ChildOfChildItem("Test1_Child1_ChildOfChild1_F1", "Test1_Child1_ChildOfChild1_F2"),
                        new ChildOfChildItem("Test1_Child1_ChildOfChild2_F1", "Test1_Child1_ChildOfChild2_F2")
                    }
                },
                new("Test1_Child2_F1", "Test1_Child2_F2"),
                new("Test1_Child3_F1", "Test1_Child3_F2")
                {
                    Children = new[]
                    {
                        new ChildOfChildItem("Test1_Child3_ChildOfChild1_F1", "Test1_Child3_ChildOfChild1_F2")
                    }
                }
            },
            ChildrenPrimitive = new[] {1}
        },
        new TestItem("Test2", new DateTime(2017, 11, 2), 110m, new Contacts("76", "753465"), Sex.Female)
        {
            ChildrenPrimitive = new[] {2, 3, 4}
        },
        new TestItem("Test3", new DateTime(2017, 11, 3), 5500.80m, new Contacts("1533", "5456"))
        {
            Children = new List<ChildItem>
            {
                new("Test3_Child1_F1", "Test3_Child1_F2")
                {
                    Children = new[]
                    {
                        new ChildOfChildItem("Test3_Child1_ChildOfChild1_F1", "Test3_Child1_ChildOfChild1_F2")
                    }
                },
                new("Test3_Child2_F1", "Test3_Child2_F2")
            },
            ChildrenPrimitive = new[] {5, 6}
        }
    };

    static DataProvider()
    {
        _conStr = Configuration.TestDbConnectionString;
    }

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
        return _testData;
    }

    public IEnumerable<ChildItem> GetChildIEnumerable(string parentName)
    {
        return _testData.SingleOrDefault(x => x.Name == parentName)?.Children;
    }

    public IEnumerable<TestItem> GetEmptyIEnumerable()
    {
        return Enumerable.Empty<TestItem>();
    }

    public int[] GetIntData()
    {
        return new[] {1, 2, 3, 4, 5, 6, 7, 8, 9, 10};
    }

    public IDataReader GetAllCustomersDataReader()
    {
        IDbConnection connection = new SqlConnection(_conStr);
        var command = connection.CreateCommand();
        command.CommandText = GetAllCustomersQuery();
        connection.Open();
        return command.ExecuteReader(CommandBehavior.CloseConnection);
    }

    public IDataReader GetEmptyDataReader()
    {
        IDbConnection connection = new SqlConnection(_conStr);
        var command = connection.CreateCommand();
        command.CommandText = GetNoCustomersQuery();
        connection.Open();
        return command.ExecuteReader(CommandBehavior.CloseConnection);
    }

    public DataTable GetAllCustomersDataTable()
    {
        var dataReader = GetAllCustomersDataReader();
        var dataTable = new DataTable();
        dataTable.Load(dataReader);
        dataReader.Close();
        return dataTable;
    }

    public DataTable GetEmptyDataTable()
    {
        var dataReader = GetEmptyDataReader();
        var dataTable = new DataTable();
        dataTable.Load(dataReader);
        dataReader.Close();
        return dataTable;
    }

    public DataSet GetAllCustomersDataSet()
    {
        using var conn = new SqlConnection(_conStr);
        conn.Open();
        var command = new SqlCommand(GetAllCustomersQuery(), conn);
        var adapter = new SqlDataAdapter(command);
        var ds = new DataSet();
        adapter.Fill(ds);
        return ds;
    }

    public DataSet GetEmptyDataSet()
    {
        using var conn = new SqlConnection(_conStr);
        conn.Open();
        var command = new SqlCommand(GetNoCustomersQuery(), conn);
        var adapter = new SqlDataAdapter(command);
        var ds = new DataSet();
        adapter.Fill(ds);
        return ds;
    }

    public IEnumerable<IDictionary<string, object>> GetDictionaryEnumerable()
    {
        return new List<IDictionary<string, object>>
        {
            new Dictionary<string, object> {["Name"] = "Name_1", ["Value"] = 25.7m, ["IsVip"] = true},
            new Dictionary<string, object> {["Name"] = "Name_2", ["Value"] = 250.7m, ["IsVip"] = false},
            new Dictionary<string, object> {["Name"] = "Name_3", ["Value"] = 2500.7m, ["IsVip"] = true}
        };
    }

    private string GetAllCustomersQuery()
    {
        return "SELECT Id, Name, IsVip, Type, Description, Revenue FROM Customers ORDER BY Id";
    }

    private string GetNoCustomersQuery()
    {
        return "SELECT Id, Name, IsVip, Type, Description, Revenue FROM Customers WHERE 1 <> 1";
    }

    public IList<ChildItem> GetChildrenProportionally(TestItem parent)
    {
        IList<ChildItem> childData = new List<ChildItem>();
        var childItemsCount = (int) parent.Sum / 2;
        childItemsCount = childItemsCount == 0 ? 1 : childItemsCount;
        for (var i = 0; i < childItemsCount; i++)
        {
            childData.Add(new ChildItem($"Parent_{parent.Sum}", $"Child_{i}"));
        }

        return childData;
    }

    public IList<ChildItem> GetChildrenRandom(int min, int max)
    {
        IList<ChildItem> childData = new List<ChildItem>();
        var childItemsCount = new Random((int) DateTime.Now.Ticks & 0x0000FFFF).Next(min, max);
        for (var i = 0; i < childItemsCount; i++)
        {
            childData.Add(new ChildItem($"Field1_{i}", $"Field2_{i}"));
        }

        return childData;
    }
}

public class TestItem
{
    public Sex Sex;

    public TestItem(string name, DateTime date, decimal sum, Contacts contacts = null, Sex sex = Sex.Male)
    {
        Name = name;
        Date = date;
        Sum = sum;
        Contacts = contacts;
        Sex = sex;
    }

    public string Name { get; set; }

    [ExcelColumn(Order = 1, AggregateFunction = AggregateFunction.Max, DisplayFormat = "yyyy-MM-dd")]
    public DateTime Date { get; set; }

    public decimal Sum { get; set; }

    public Contacts Contacts { get; set; }

    public IEnumerable<ChildItem> Children { get; set; }

    public IEnumerable<int> ChildrenPrimitive { get; set; }
}

public enum Sex
{
    Male,
    Female
}

public class ChildItem
{
    public ChildItem(string field1, string field2)
    {
        Field1 = field1;
        Field2 = field2;
    }

    public string Field1 { get; set; }

    [ExcelColumn(AggregateFunction = AggregateFunction.Min)]
    public string Field2 { get; set; }

    public ChildOfChildItem[] Children { get; set; }
}

public class ChildOfChildItem
{
    public ChildOfChildItem(string field1, string field2)
    {
        Field1 = field1;
        Field2 = field2;
    }

    public string Field1 { get; set; }

    public string Field2 { get; set; }
}

public class Contacts
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