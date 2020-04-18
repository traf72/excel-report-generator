using ExcelReportGenerator.Rendering.Providers.VariableProviders;
using System;

namespace ExcelReportGenerator.Samples.Customizations
{
    public class CustomSystemVariableProvider : SystemVariableProvider
    {
        public string ReportTime => DateTime.Now.ToString("g");
    }
}