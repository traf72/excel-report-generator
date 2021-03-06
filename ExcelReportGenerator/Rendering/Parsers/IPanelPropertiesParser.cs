﻿using System.Collections.Generic;

namespace ExcelReportGenerator.Rendering.Parsers
{
    internal interface IPanelPropertiesParser
    {
        IDictionary<string, string> Parse(string input);
    }
}