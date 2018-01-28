using System.Collections;

namespace ExcelReportGenerator.Enumerators
{
    internal interface ICustomEnumerator : IEnumerator
    {
        int RowCount { get; }
    }
}