using ExcelReportGenerator.Helpers;
using System.Collections;
using System.Data;

namespace ExcelReportGenerator.Enumerators;

internal class DataSetEnumerator : IGenericCustomEnumerator<DataRow>
{
    private readonly IGenericCustomEnumerator<DataRow> _dataTableEnumerator;

    public DataSetEnumerator(DataSet dataSet, string tableName = null)
    {
        if (dataSet == null)
        {
            throw new ArgumentNullException(nameof(dataSet), ArgumentHelper.NullParamMessage);
        }
        if (dataSet.Tables.Count == 0)
        {
            throw new InvalidOperationException("DataSet does not contain any table");
        }

        DataTable dataTable;
        if (!string.IsNullOrWhiteSpace(tableName))
        {
            dataTable = dataSet.Tables[tableName];
            if (dataTable == null)
            {
                throw new InvalidOperationException($"DataSet does not contain table with name \"{tableName}\"");
            }
        }
        else
        {
            dataTable = dataSet.Tables[0];
        }

        _dataTableEnumerator = new DataTableEnumerator(dataTable);
    }

    public DataRow Current => _dataTableEnumerator.Current;

    object IEnumerator.Current => Current;

    public bool MoveNext() => _dataTableEnumerator.MoveNext();

    public void Reset() => _dataTableEnumerator.Reset();

    public int RowCount => _dataTableEnumerator.RowCount;

    public void Dispose() => _dataTableEnumerator.Dispose();
}