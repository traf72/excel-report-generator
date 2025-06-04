using System.Globalization;

namespace ExcelReportGenerator.Tests
{
    [SetUpFixture]
    public class GlobalSetup
    {
        [OneTimeSetUp]
        public void SetUp()
        {
            var culture = new CultureInfo("ru-RU");

            CultureInfo.CurrentCulture = culture;
            CultureInfo.CurrentUICulture = culture;

            Thread.CurrentThread.CurrentCulture = culture;
            Thread.CurrentThread.CurrentUICulture = culture;
        }
    }
}