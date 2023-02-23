using Microsoft.Extensions.Configuration;

namespace ExcelReportGenerator.Tests;

public static class Configuration
{
    private static readonly IConfigurationRoot _configuration;

    static Configuration()
    {
        var builder = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", true, true);
        _configuration = builder.Build();
    }

    public static string TestDbConnectionString
    {
        get
        {
            var projectPath = new DirectoryInfo(Directory.GetCurrentDirectory()).Parent.Parent.Parent.FullName;
            return _configuration["ConnectionStrings:TestDb"].Replace("%DBPATH%", projectPath);
        }
    }
}