using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace ExcelReportGenerator.Samples;

public static class ConnectionFactory
{
    public static IDbConnection Create()
    {
        var connection = new SqlConnection(ConfigurationManager.ConnectionStrings["AdventureWorks"].ConnectionString);
        connection.Open();
        return connection;
    }
}