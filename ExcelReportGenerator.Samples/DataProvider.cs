using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using ExcelReportGenerator.Attributes;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Samples.Extensions;

namespace ExcelReportGenerator.Samples
{
    public class DataProvider
    {
        public IDataReader GetEmployeesAsIDataReader(string department = null)
        {
            var con = ConnectionFactory.Create();
            using var command = con.CreateCommand();
            command.CommandText = Query;
            command.Parameters.Add(new SqlParameter("@department", string.IsNullOrWhiteSpace(department) ? (object)DBNull.Value : department));
            return command.ExecuteReader(CommandBehavior.CloseConnection);
        }

        public DataTable GetEmployeesAsDataTable(string department = null)
        {
            using var reader = GetEmployeesAsIDataReader(department);
            var dataTable = new DataTable();
            dataTable.Load(reader);
            return dataTable;
        }

        public DataSet GetEmployeesAsDataSet(string department = null)
        {
            using var con = (SqlConnection)ConnectionFactory.Create();
            using var command = new SqlCommand(Query, con);
            command.Parameters.Add(new SqlParameter("@department", string.IsNullOrWhiteSpace(department) ? (object)DBNull.Value : department));
            var adapter = new SqlDataAdapter(command);
            var dataSet = new DataSet();
            adapter.Fill(dataSet);
            return dataSet;
        }

        public IEnumerable<Result> GetEmployeesAsIEnumerable(string department = null)
        {
            IList<Result> result = new List<Result>();
            using (var reader = GetEmployeesAsIDataReader(department))
            {
                while (reader.Read())
                {
                    result.Add(new Result
                    {
                        LastName = reader.GetString(reader.GetOrdinal("LastName")),
                        FirstName = reader.GetString(reader.GetOrdinal("FirstName")),
                        MiddleName = reader.GetValueSafe<string>(reader.GetOrdinal("MiddleName")),
                        JobTitle = reader.GetString(reader.GetOrdinal("JobTitle")),
                        BirthDate = reader.GetDateTime(reader.GetOrdinal("BirthDate")),
                        Sex = reader.GetString(reader.GetOrdinal("Gender")),
                        HireDate = reader.GetDateTime(reader.GetOrdinal("HireDate")),
                        Rate = reader.GetDecimal(reader.GetOrdinal("Rate")),
                        DepartmentName = reader.GetString(reader.GetOrdinal("DepartmentName")),
                    });
                }
            }

            return result;
        }

        public IEnumerable<IDictionary<string, object>> GetEmployeesAsIEnumerableOfDictionary(string department = null)
        {
            IList<IDictionary<string, object>> result = new List<IDictionary<string, object>>();
            using (var reader = GetEmployeesAsIDataReader(department))
            {
                while (reader.Read())
                {
                    IDictionary<string, object> dict = new Dictionary<string, object>();

                    dict["LastName"] = reader.GetString(reader.GetOrdinal("LastName"));
                    dict["FirstName"] = reader.GetString(reader.GetOrdinal("FirstName"));
                    dict["MiddleName"] = reader.GetValueSafe<string>(reader.GetOrdinal("MiddleName"));
                    dict["JobTitle"] = reader.GetString(reader.GetOrdinal("JobTitle"));
                    dict["BirthDate"] = reader.GetDateTime(reader.GetOrdinal("BirthDate"));
                    dict["Gender"] = reader.GetString(reader.GetOrdinal("Gender"));
                    dict["HireDate"] = reader.GetDateTime(reader.GetOrdinal("HireDate"));
                    dict["Rate"] = reader.GetDecimal(reader.GetOrdinal("Rate"));
                    dict["DepartmentName"] = reader.GetString(reader.GetOrdinal("DepartmentName"));

                    result.Add(dict);
                }
            }

            return result;
        }

        private static string Query =>
            @"
                    SELECT
	                    p.LastName
	                    , p.FirstName
	                    , p.MiddleName
	                    , e.BusinessEntityID
	                    , e.JobTitle
	                    , e.BirthDate
	                    , e.Gender
	                    , e.HireDate
	                    , t.Rate
	                    , d.Name AS DepartmentName
                    FROM
	                    HumanResources.Employee e
	                    JOIN HumanResources.EmployeeDepartmentHistory dh ON e.BusinessEntityId = dh.BusinessEntityId
	                    JOIN HumanResources.Department d ON d.DepartmentId = dh.DepartmentId
	                    CROSS APPLY
	                    (
		                    SELECT TOP 1
			                    Rate
		                    FROM
			                    HumanResources.EmployeePayHistory
		                    WHERE
			                    BusinessEntityID = e.BusinessEntityID
		                    ORDER BY
			                    RateChangeDate DESC
	                    ) t
	                    JOIN Person.BusinessEntity be ON e.BusinessEntityID = be.BusinessEntityID
	                    JOIN Person.Person p ON be.BusinessEntityID = p.BusinessEntityID
                    WHERE
	                    dh.EndDate IS NULL
	                    AND d.Name = ISNULL(@department, d.Name)
                    ORDER BY
                        d.Name
	                    , p.LastName
	                    , p.FirstName
	                    , p.MiddleName
                    ";

        public class Result
        {
            public string DepartmentName { get; set; }

            public string LastName { get; set; }

            public string FirstName { get; set; }

            [NullValue("N/A")]
            public string MiddleName { get; set; }

            [ExcelColumn(AdjustToContent = true)]
            public string JobTitle { get; set; }

            [ExcelColumn(Caption = "Gender")]
            public string Sex { get; set; }

            [ExcelColumn(AggregateFunction = AggregateFunction.Max)]
            public DateTime BirthDate { get; set; }

            [ExcelColumn(AggregateFunction = AggregateFunction.Min)]
            public DateTime HireDate { get; set; }

            [ExcelColumn(DisplayFormat = "$#,0.00")]
            public decimal Rate { get; set; }
        }
    }
}