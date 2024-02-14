using ExcelReportGenerator.Extensions;
using Assert = NUnit.Framework.Legacy.ClassicAssert;

namespace ExcelReportGenerator.Tests.Extensions;

public class StringExtensionsTest
{
    [Test]
    public void TestReplaceFirst()
    {
        var input = "Hello World";
        Assert.AreEqual("Helle World", input.ReplaceFirst("o", "e"));
        Assert.AreEqual("Hello World", input.ReplaceFirst("O", "e"));
        Assert.AreEqual("ppello World", input.ReplaceFirst("H", "pp"));
        Assert.AreEqual("HeLLlo World", input.ReplaceFirst("l", "LL"));
        Assert.IsNull(StringExtensions.ReplaceFirst(null, "a", "A"));
    }

    [Test]
    public void TestReverse()
    {
        Assert.AreEqual("!dlroW olleH", "Hello World!".Reverse());
        Assert.AreEqual(string.Empty, string.Empty.Reverse());
        Assert.IsNull(StringExtensions.Reverse(null));
    }
}