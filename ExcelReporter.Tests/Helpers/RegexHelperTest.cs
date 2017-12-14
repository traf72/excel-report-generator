using ExcelReporter.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Helpers
{
    [TestClass]
    public class RegexHelperTest
    {
        [TestMethod]
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
}