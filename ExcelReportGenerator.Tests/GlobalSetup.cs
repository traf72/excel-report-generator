using System.Globalization;

namespace ExcelReportGenerator.Tests
{
    [SetUpFixture]
    public class GlobalSetup
    {
        [OneTimeSetUp]
        public void SetUp()
        {
            var culture = new CultureInfo("ka-GE");

            CultureInfo.CurrentCulture = culture;
            CultureInfo.CurrentUICulture = culture;

            culture.DateTimeFormat.ShortTimePattern = "H:mm:ss";
            culture.DateTimeFormat.LongTimePattern = "H:mm:ss";

            Thread.CurrentThread.CurrentCulture = culture;
            Thread.CurrentThread.CurrentUICulture = culture;
        }
    }
}