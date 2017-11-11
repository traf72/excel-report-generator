using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;

namespace ExcelReporter.Enumerators
{
    internal class DataReaderEnumerator : IEnumerator<DataRow>
    {
        private const string ClosedDataReaderMessage = "DataReader has been closed";
        private const string FinishedEnumeratorMessage = "Enumerator has been finished";

        private readonly IDataReader _dataReader;
        private DataTable _dataTable;
        private DataRow _currentDataRow;
        private bool _isStarted;
        private bool _isFinished;

        public DataReaderEnumerator(IDataReader dataReader)
        {
            _dataReader = dataReader ?? throw new ArgumentNullException(nameof(dataReader), Constants.NullParamMessage);
            CreateDataTable();
        }

        private void CreateDataTable()
        {
            DataTable schemaTable = _dataReader.GetSchemaTable();
            if (schemaTable != null)
            {
                _dataTable = new DataTable();
                foreach (DataRow row in schemaTable.Rows)
                {
                    string colName = row.Field<string>("ColumnName");
                    Type type = row.Field<Type>("DataType");
                    _dataTable.Columns.Add(colName, type);
                }
            }
        }

        public DataRow Current
        {
            get
            {
                if (_dataReader.IsClosed)
                {
                    throw new InvalidOperationException(ClosedDataReaderMessage);
                }
                if (!_isStarted)
                {
                    throw new InvalidOperationException("Enumerator has not been started. Call MoveNext() method.");
                }
                if (_isFinished)
                {
                    throw new InvalidOperationException(FinishedEnumeratorMessage);
                }
                return _currentDataRow;
            }
        }

        object IEnumerator.Current => Current;

        public bool MoveNext()
        {
            if (_dataReader.IsClosed)
            {
                throw new InvalidOperationException(ClosedDataReaderMessage);
            }
            if (_isFinished)
            {
                throw new InvalidOperationException(FinishedEnumeratorMessage);
            }
            _isStarted = true;
            _isFinished = !_dataReader.Read();
            if (!_isFinished)
            {
                ExtractDataRow();
            }

            return !_isFinished;
        }

        private void ExtractDataRow()
        {
            if (_dataTable == null)
            {
                return;
            }

            _currentDataRow = _dataTable.Rows.Add();
            foreach (DataColumn col in _dataTable.Columns)
            {
                _currentDataRow[col.ColumnName] = _dataReader[col.ColumnName];
            }
        }

        public void Reset() => throw new NotSupportedException($"{nameof(DataReaderEnumerator)} does not support reset method");

        public void Dispose() => _dataReader.Close();
    }
}