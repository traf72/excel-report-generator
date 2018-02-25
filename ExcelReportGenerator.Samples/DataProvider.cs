using ExcelReportGenerator.Samples.Extensions;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace ExcelReportGenerator.Samples
{
    public class DataProvider
    {
        public IDataReader GetEmployeesAsIDataReader(string department = null)
        {
            const string query = @"
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

            IDbConnection con = ConnectionFactory.Create();
            using (IDbCommand command = con.CreateCommand())
            {
                command.CommandText = query;
                command.Parameters.Add(new SqlParameter("@department", string.IsNullOrWhiteSpace(department) ? (object) DBNull.Value : department));
                return command.ExecuteReader(CommandBehavior.CloseConnection);
            }
        }

        public DataTable GetEmployeesAsDataTable(string department)
        {
            using (IDataReader reader = GetEmployeesAsIDataReader(department))
            {
                var dataTable = new DataTable();
                dataTable.Load(reader);
                return dataTable;
            }
        }

        public IEnumerable<Result> GetEmployeesAsIEnumerable(string department)
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
                        Gender = reader.GetChar(reader.GetOrdinal("Gender")),
                        HireDate = reader.GetDateTime(reader.GetOrdinal("HireDate")),
                        Rate = reader.GetDecimal(reader.GetOrdinal("Rate")),
                        DepartmentName = reader.GetString(reader.GetOrdinal("DepartmentName")),
                    });
                }
            }

            return result;
        }

        public class Result
        {
            public string LastName { get; set; }
            public string FirstName { get; set; }
            public string MiddleName { get; set; }
            public string JobTitle { get; set; }
            public DateTime BirthDate { get; set; }
            public char Gender { get; set; }
            public DateTime HireDate { get; set; }
            public decimal Rate { get; set; }
            public string DepartmentName { get; set; }
        }
    }
}