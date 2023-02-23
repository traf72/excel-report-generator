using ExcelReportGenerator.Helpers;
using System.Collections;
using System.Data;

namespace ExcelReportGenerator.Enumerators;

internal class DataReaderEnumerator : IGenericCustomEnumerator<DataRow>
{
    private readonly IGenericCustomEnumerator<DataRow> _dataTableEnumerator;

    public DataReaderEnumerator(IDataReader dataReader)
    {
        if (dataReader == null)
        {
            throw new ArgumentNullException(nameof(dataReader), ArgumentHelper.NullParamMessage);
        }

        var dataTable = new DataTable();
        dataTable.Load(dataReader);
        dataReader.Dispose();
        _dataTableEnumerator = new DataTableEnumerator(dataTable);
    }

    public DataRow Current => _dataTableEnumerator.Current;

    object IEnumerator.Current => Current;

    public bool MoveNext() => _dataTableEnumerator.MoveNext();

    public void Reset() => _dataTableEnumerator.Reset();

    public int RowCount => _dataTableEnumerator.RowCount;

    public void Dispose() => _dataTableEnumerator.Dispose();
}