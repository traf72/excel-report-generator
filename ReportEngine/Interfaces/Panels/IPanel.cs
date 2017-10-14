using ClosedXML.Excel;
using ReportEngine.Enums;
using ReportEngine.Interfaces.Reports;
using System.Collections.Generic;
using JetBrains.Annotations;

namespace ReportEngine.Interfaces.Panels
{
    public interface IPanel
    {
        void Render();

        IPanel Parent { get; set; }

        IEnumerable<IPanel> Children { get; set; }

        IExcelReport Context { get; set; }

        IXLRange Range { get; }

        ShiftType ShiftType { get; set; }

        PanelType Type { get; set; }

        int RenderPriority { get; set; }

        IPanel Copy([NotNull] IXLCell cell, bool recursive = true);

        void Move([NotNull] IXLCell cell);

        /// <summary>
        /// Пересчитывает Range относительно родительского, а также Range'и всех Children'ов
        /// Только для внутренних целей
        /// </summary>
        void RecalculateRangeRelativeParentRecursive();

        void Delete();
    }
}