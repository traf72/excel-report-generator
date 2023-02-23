namespace ExcelReportGenerator.Rendering.Panels;

internal interface IPanel
{
    void Delete();

    string BeforeRenderMethodName { get; set; }

    string AfterRenderMethodName { get; set; }
}