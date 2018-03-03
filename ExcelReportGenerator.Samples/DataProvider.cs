using ExcelReportGenerator.Samples.Extensions;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Samples
{
    public class DataProvider
    {
        public IDataReader GetEmployeesAsIDataReader(string department = null)
        {
            IDbConnection con = ConnectionFactory.Create();
            using (IDbCommand command = con.CreateCommand())
            {
                command.CommandText = Query;
                command.Parameters.Add(new SqlParameter("@department", string.IsNullOrWhiteSpace(department) ? (object)DBNull.Value : department));
                return command.ExecuteReader(CommandBehavior.CloseConnection);
            }
        }

        public DataTable GetEmployeesAsDataTable(string department = null)
        {
            using (IDataReader reader = GetEmployeesAsIDataReader(department))
            {
                var dataTable = new DataTable();
                dataTable.Load(reader);
                return dataTable;
            }
        }

        public DataSet GetEmployeesAsDataSet(string department = null)
        {
            using (var con = (SqlConnection)ConnectionFactory.Create())
            {
                using (var command = new SqlCommand(Query, con))
                {
                    command.Parameters.Add(new SqlParameter("@department", string.IsNullOrWhiteSpace(department) ? (object)DBNull.Value : department));
                    var adapter = new SqlDataAdapter(command);
                    var dataSet = new DataSet();
                    adapter.Fill(dataSet);
                    return dataSet;
                }
            }
        }

        public IEnumerable<Result> GetEmployeesAsIEnumerable(string department = null)
        {
            IList<Result> result = new List<Result>();
            using (IDataReader reader = GetEmployeesAsIDataReader(department))
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
                        Gender = reader.GetString(reader.GetOrdinal("Gender")),
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
            using (IDataReader reader = GetEmployeesAsIDataReader(department))
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

        private string Query
        {
            get
            {
                return @"
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
            }
        }

        public class Result
        {
            public string LastName { get; set; }

            public string FirstName { get; set; }

            [NullValue("N/A")]
            public string MiddleName { get; set; }

            public string JobTitle { get; set; }

            public DateTime BirthDate { get; set; }

            public string Gender { get; set; }

            public DateTime HireDate { get; set; }

            public decimal Rate { get; set; }

            public string DepartmentName { get; set; }
        }
    }
}