using ClosedXML.Excel;

namespace ExcelReporter.Interfaces.Panels.Excel
{
    public interface IExcelNamedPanel : INamedPanel, IExcelPanel
    {
        IExcelNamedPanel Copy(IXLCell cell, string name, bool recursive = true);

        void RemoveName(bool recursive = false);
    }
}