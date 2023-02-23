using ExcelReportGenerator.Helpers;

namespace ExcelReportGenerator.Tests.Helpers;

public class RegexHelperTest
{
    [Test]
    public void TestSafeEscape()
    {
        Assert.IsNull(RegexHelper.SafeEscape(null));
        Assert.AreEqual(string.Empty, RegexHelper.SafeEscape(string.Empty));
        Assert.AreEqual("abc", RegexHelper.SafeEscape("abc"));
        Assert.AreEqual("\\ ", RegexHelper.SafeEscape(" "));
        Assert.AreEqual("\\[", RegexHelper.SafeEscape("["));
        Assert.AreEqual("]", RegexHelper.SafeEscape("]"));
        Assert.AreEqual("\\*\\*", RegexHelper.SafeEscape("**"));
    }
}