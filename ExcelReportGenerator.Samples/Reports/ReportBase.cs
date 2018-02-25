namespace ExcelReportGenerator.Samples.Reports
{
    public abstract class ReportBase
    {
        public abstract string ReportName { get; }

        public string ConvertGender(string gender)
        {
            return gender == "M" ? "Male" : "Female";
        }
    }
}