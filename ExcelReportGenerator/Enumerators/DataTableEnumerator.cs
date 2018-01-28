using ExcelReportGenerator.Helpers;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;

namespace ExcelReportGenerator.Enumerators
{
    internal class DataTableEnumerator : IGenericCustomEnumerator<DataRow>
    {
        private readonly DataTable _dataTable;

        private IEnumerator<DataRow> _dataTableEnumerator;

        public DataTableEnumerator(DataTable dataTable)
        {
            _dataTable = dataTable ?? throw new ArgumentNullException(nameof(dataTable), ArgumentHelper.NullParamMessage);
            _dataTableEnumerator = _dataTable.AsEnumerable().GetEnumerator();
        }

        public DataRow Current => _dataTableEnumerator.Current;

        object IEnumerator.Current => Current;

        public bool MoveNext() => _dataTableEnumerator.MoveNext();

        public void Reset() => _dataTableEnumerator = _dataTable.AsEnumerable().GetEnumerator();

        public int RowCount => _dataTable.Rows.Count;

        public void Dispose() => _dataTableEnumerator.Dispose();
    }
}