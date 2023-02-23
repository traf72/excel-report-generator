using ExcelReportGenerator.Rendering;

namespace ExcelReportGenerator.Samples.Customizations;

public class CustomSystemFunctions : SystemFunctions
{
    public static string ConvertGender(string gender) => gender == "M" ? "Male" : "Female";
}