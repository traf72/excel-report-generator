using System.Collections.Generic;

namespace ExcelReportGenerator.Enumerators
{
    public interface ICustomEnumerator<out T> : IEnumerator<T>
    {
        int RowCount { get; }
    }
}