using System;
using System.Data;

namespace ExcelReporter.Iterators
{
    public class DataSetIterator : IIterator<DataRow>
    {
        private readonly DataTableIterator _dataTableIterator;

        public DataSetIterator(DataSet dataSet, string tableName = null)
        {
            if (dataSet == null)
            {
                throw new ArgumentNullException(nameof(dataSet), Constants.NullParamMessage);
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

            _dataTableIterator = new DataTableIterator(dataTable);
        }

        public DataRow Next()
        {
            return _dataTableIterator.Next();
        }

        public bool HaxNext()
        {
            return _dataTableIterator.HaxNext();
        }

        public void Reset()
        {
            _dataTableIterator.Reset();
        }
    }
}