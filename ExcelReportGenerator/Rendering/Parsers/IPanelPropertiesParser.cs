using System.Collections.Generic;

namespace ExcelReportGenerator.Rendering.Parsers
{
    public interface IPanelPropertiesParser
    {
        IDictionary<string, string> Parse(string input);
    }
}