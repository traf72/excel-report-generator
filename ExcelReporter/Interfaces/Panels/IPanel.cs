using ClosedXML.Excel;
using ExcelReporter.Enums;
using ExcelReporter.Interfaces.Reports;
using System.Collections.Generic;

namespace ExcelReporter.Interfaces.Panels
{
    public interface IPanel
    {
        void Render();

        IPanel Parent { get; set; }

        IEnumerable<IPanel> Children { get; set; }

        IExcelReport Report { get; set; }

        IXLRange Range { get; }

        ShiftType ShiftType { get; set; }

        PanelType Type { get; set; }

        int RenderPriority { get; set; }

        IPanel Copy(IXLCell cell, bool recursive = true);

        void Move(IXLCell cell);

        /// <summary>
        /// Пересчитывает Range относительно родительского, а также Range'и всех Children'ов
        /// Только для внутренних целей
        /// </summary>
        void RecalculateRangeRelativeParentRecursive();

        void Delete();
    }
}