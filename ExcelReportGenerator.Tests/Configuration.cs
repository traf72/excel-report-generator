using Microsoft.Extensions.Configuration;
using System.IO;

namespace ExcelReportGenerator.Tests
{
    public static class Configuration
    {
        private static readonly IConfigurationRoot _configuration;

        static Configuration()
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
            _configuration = builder.Build();
        }

        public static string TestDbConnectionString
        {
            get
            {
                string projectPath = new DirectoryInfo(Directory.GetCurrentDirectory()).Parent.Parent.Parent.FullName;
                return _configuration["ConnectionStrings:TestDb"].Replace("%DBPATH%", projectPath);
            }
        }
    }
}