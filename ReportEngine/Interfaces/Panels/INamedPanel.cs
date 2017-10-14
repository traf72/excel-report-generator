using ClosedXML.Excel;
using JetBrains.Annotations;

namespace ReportEngine.Interfaces.Panels
{
    public interface INamedPanel : IPanel
    {
        string Name { get; }

        INamedPanel Copy([NotNull] IXLCell cell, [NotNull] string name, bool recursive = true);

        void RemoveName(bool recursive = false);
    }
}