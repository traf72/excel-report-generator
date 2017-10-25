using ExcelReporter.Exceptions;
using ExcelReporter.Interfaces.Providers;
using System;
using System.Data;
using ExcelReporter.Interfaces.Providers.DataItemValueProviders;

namespace ExcelReporter.Implementations.Providers.DataItemValueProviders
{
    public class DataRowValueProvider : IGenericDataItemValueProvider<DataRow>
    {
        private string _columnName;
        private DataRow _dataRow;

        public virtual object GetValue(string columnName, DataRow dataRow)
        {
            if (string.IsNullOrWhiteSpace(columnName))
            {
                throw new ArgumentException(Constants.EmptyStringParamMessage, nameof(columnName));
            }
            if (dataRow == null)
            {
                throw new ArgumentNullException(nameof(dataRow), Constants.NullParamMessage);
            }

            _columnName = columnName.Trim();
            _dataRow = dataRow;
            return dataRow.ItemArray[GetColumnIndex()];
        }

        private int GetColumnIndex()
        {
            DataColumn column = _dataRow.Table.Columns[_columnName];
            if (column == null)
            {
                throw new ColumnNotFoundException($"DataRow does not contain column \"{_columnName}\"");
            }
            return column.Ordinal;
        }

        object IDataItemValueProvider.GetValue(string columnName, object dataRow)
        {
            return GetValue(columnName, (DataRow)dataRow);
        }
    }
}