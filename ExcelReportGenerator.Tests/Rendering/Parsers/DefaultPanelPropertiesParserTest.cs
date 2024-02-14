using ExcelReportGenerator.Rendering;
using ExcelReportGenerator.Rendering.Parsers;
using Assert = NUnit.Framework.Legacy.ClassicAssert;

namespace ExcelReportGenerator.Tests.Rendering.Parsers;

public class DefaultPanelPropertiesParserTest
{
    [Test]
    public void TestParse()
    {
        var input =
            $"Prop1=Val1;Prop2 = Val2 {Environment.NewLine} Prop6 \t\t \t Prop4=Val4; Prop5<Val5{Environment.NewLine}Prop3=Val3";
        IPanelPropertiesParser parser = new DefaultPanelPropertiesParser(new PanelParsingSettings
        {
            PanelPropertiesSeparators = new[] {Environment.NewLine, "\t", ";"},
            PanelPropertyNameValueSeparator = "="
        });

        var result = parser.Parse(input);
        Assert.AreEqual(4, result.Count);
        Assert.AreEqual(result["Prop1"], "Val1");
        Assert.AreEqual(result["Prop2"], "Val2");
        Assert.AreEqual(result["Prop3"], "Val3");
        Assert.AreEqual(result["Prop4"], "Val4");

        input =
            "Prop1==Val1&&Prop2 == Val2 ;; Prop6 ;;&& Prop4=Val4;; Prop5<Val5&&Prop3==Val3;;Prop2==Override  && prop1==val1";
        parser = new DefaultPanelPropertiesParser(new PanelParsingSettings
        {
            PanelPropertiesSeparators = new[] {";;", "&&"},
            PanelPropertyNameValueSeparator = "=="
        });

        result = parser.Parse(input);
        Assert.AreEqual(4, result.Count);
        Assert.AreEqual(result["Prop1"], "Val1");
        Assert.AreEqual(result["Prop2"], "Override");
        Assert.AreEqual(result["Prop3"], "Val3");
        Assert.AreEqual(result["prop1"], "val1");
    }

    [Test]
    public void TestParseIfInputEmpty()
    {
        string input = null;
        IPanelPropertiesParser parser = new DefaultPanelPropertiesParser(new PanelParsingSettings
        {
            PanelPropertiesSeparators = new[] {Environment.NewLine, "\t", ";"},
            PanelPropertyNameValueSeparator = "="
        });

        var result = parser.Parse(input);
        Assert.AreEqual(0, result.Count);

        input = string.Empty;

        result = parser.Parse(input);
        Assert.AreEqual(0, result.Count);

        input = " ";

        result = parser.Parse(input);
        Assert.AreEqual(0, result.Count);
    }
}