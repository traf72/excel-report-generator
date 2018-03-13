using ExcelReportGenerator.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReportGenerator.Tests.Extensions
{
    [TestClass]
    public class StringExtensionsTest
    {
        [TestMethod]
        public void TestReplaceFirst()
        {
            string input = "Hello World";
            Assert.AreEqual("Helle World", input.ReplaceFirst("o", "e"));
            Assert.AreEqual("Hello World", input.ReplaceFirst("O", "e"));
            Assert.AreEqual("ppello World", input.ReplaceFirst("H", "pp"));
            Assert.AreEqual("HeLLlo World", input.ReplaceFirst("l", "LL"));
            Assert.IsNull(StringExtensions.ReplaceFirst(null, "a", "A"));
        }

        [TestMethod]
        public void TestReverse()
        {
            Assert.AreEqual("!dlroW olleH", "Hello World!".Reverse());
            Assert.AreEqual(string.Empty, string.Empty.Reverse());
            Assert.IsNull(StringExtensions.Reverse(null));
        }
    }
}