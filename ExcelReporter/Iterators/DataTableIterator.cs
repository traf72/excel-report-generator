using System;
using System.Data;

namespace ExcelReporter.Iterators
{
    public class DataTableIterator : IIterator<DataRow>
    {
        private readonly DataTable _dataTable;
        private int _currentRowIndex = -1;

        public DataTableIterator(DataTable dataTable)
        {
            if (dataTable == null)
            {
                throw new ArgumentNullException(nameof(dataTable), Constants.NullParamMessage);
            }
            _dataTable = dataTable;
        }

        public DataRow Next()
        {
            if (_currentRowIndex == -1)
            {
                throw new InvalidOperationException("Iterator has not been started");
            }
            if (_currentRowIndex >= _dataTable.Rows.Count)
            {
                throw new InvalidOperationException("Iterator has been finished");
            }

            return _dataTable.Rows[_currentRowIndex];
        }

        public bool HaxNext()
        {
            bool result = _dataTable.Rows.Count > _currentRowIndex + 1;
            _currentRowIndex++;
            return result;
        }

        public void Reset()
        {
            _currentRowIndex = -1;
        }
    }
}