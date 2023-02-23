using ExcelReportGenerator.Rendering.Providers.VariableProviders;

namespace ExcelReportGenerator.Samples.Customizations;

public class CustomSystemVariableProvider : SystemVariableProvider
{
    public string ReportTime => DateTime.Now.ToString("g");
}