using System;
using System.Collections.Generic;

namespace ExcelReportGenerator.Rendering.Parsers
{
    public class DefaultPanelPropertiesParser : IPanelPropertiesParser
    {
        private readonly PanelParsingSettings _panelParsingSettings;

        public DefaultPanelPropertiesParser(PanelParsingSettings panelParsingSettings)
        {
            _panelParsingSettings = panelParsingSettings;
        }

        public IDictionary<string, string> Parse(string input)
        {
            IDictionary<string, string> result = new Dictionary<string, string>();
            if (string.IsNullOrWhiteSpace(input))
            {
                return result;
            }

            string[] props = input.Split(_panelParsingSettings.PanelPropertiesSeparators, StringSplitOptions.RemoveEmptyEntries);
            foreach (string prop in props)
            {
                int propNameSeparatorIndex = prop.IndexOf(_panelParsingSettings.PanelPropertyNameValueSeparator, StringComparison.Ordinal);
                if (propNameSeparatorIndex == -1)
                {
                    continue;
                }

                string name = prop.Substring(0, propNameSeparatorIndex).Trim();
                string value = prop.Substring(propNameSeparatorIndex + _panelParsingSettings.PanelPropertyNameValueSeparator.Length).Trim();
                result[name] = value;
            }

            return result;
        }
    }
}