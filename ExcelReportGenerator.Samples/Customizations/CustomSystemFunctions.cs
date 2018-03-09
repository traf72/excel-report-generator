using ExcelReportGenerator.Rendering;

namespace ExcelReportGenerator.Samples.Customizations
{
    public class CustomSystemFunctions : SystemFunctions
    {
        public static string ConvertGender(string gender)
        {
            return gender == "M" ? "Male" : "Female";
        }
    }
}