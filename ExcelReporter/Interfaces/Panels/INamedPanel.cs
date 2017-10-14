using ClosedXML.Excel;

namespace ExcelReporter.Interfaces.Panels
{
    public interface INamedPanel : IPanel
    {
        string Name { get; }

        INamedPanel Copy(IXLCell cell, string name, bool recursive = true);

        void RemoveName(bool recursive = false);
    }
}