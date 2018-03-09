using ExcelReportGenerator.Samples.Customizations;

namespace ExcelReportGenerator.Samples.Reports
{
    public abstract class ReportBase
    {
        public abstract string ReportName { get; }

        public string ConvertGender(string gender)
        {
            return CustomSystemFunctions.ConvertGender(gender);
        }
    }
}