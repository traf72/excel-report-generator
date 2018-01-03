using System.Collections.Generic;

namespace ExcelReporter.Rendering.Parsers
{
    public interface IPanelPropertiesParser
    {
        IDictionary<string, string> Parse(string input);
    }
}