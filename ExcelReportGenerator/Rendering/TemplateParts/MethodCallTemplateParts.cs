using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Rendering.TemplateParts
{
    /// <summary>
    /// Represent parts from which method template consist of
    /// </summary>
    [LicenceKeyPart(L = true, U = true)]
    public class MethodCallTemplateParts : MemberTemplateParts
    {
        public MethodCallTemplateParts(string typeName, string memberName, string methodParams) : base(typeName, memberName)
        {
            MethodParams = methodParams;
        }

        public string MethodParams { get; }
    }
}