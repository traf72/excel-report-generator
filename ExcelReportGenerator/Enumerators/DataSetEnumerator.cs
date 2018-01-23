using ExcelReportGenerator.Helpers;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;

namespace ExcelReportGenerator.Enumerators
{
    internal class DataSetEnumerator : ICustomEnumerator<DataRow>
    {
        private readonly DataTable _dataTable;

        private readonly IEnumerator<DataRow> _dataTableEnumerator;

        public DataSetEnumerator(DataSet dataSet, string tableName = null)
        {
            _ = dataSet ?? throw new ArgumentNullException(nameof(dataSet), ArgumentHelper.NullParamMessage);
            if (dataSet.Tables.Count == 0)
            {
                throw new InvalidOperationException("DataSet does not contain any table");
            }

            if (!string.IsNullOrWhiteSpace(tableName))
            {
                _dataTable = dataSet.Tables[tableName];
                if (_dataTable == null)
                {
                    throw new InvalidOperationException($"DataSet does not contain table with name \"{tableName}\"");
                }
            }
            else
            {
                _dataTable = dataSet.Tables[0];
            }

            _dataTableEnumerator = _dataTable.AsEnumerable().GetEnumerator();
        }

        public DataRow Current => _dataTableEnumerator.Current;

        object IEnumerator.Current => Current;

        public bool MoveNext() => _dataTableEnumerator.MoveNext();

        public void Reset() => _dataTableEnumerator.Reset();

        public void Dispose() => _dataTableEnumerator.Dispose();

        public int RowCount => _dataTable.Rows.Count;
    }
}