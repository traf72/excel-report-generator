using ClosedXML.Excel;

namespace ExcelReporter.Rendering.Panels.ExcelPanels
{
    internal interface IExcelNamedPanel : INamedPanel, IExcelPanel
    {
        IExcelNamedPanel Copy(IXLCell cell, string name, bool recursive = true);

        void RemoveName(bool recursive = false);
    }
}