using ExcelReporter.Exceptions;
using ExcelReporter.Interfaces.Providers.DataItemValueProviders;
using System;
using System.Data;

namespace ExcelReporter.Implementations.Providers.DataItemValueProviders
{
    /// <summary>
    /// Provides values from data row
    /// </summary>
    public class DataRowValueProvider : IGenericDataItemValueProvider<DataRow>
    {
        private string _columnName;
        private DataRow _dataRow;

        /// <summary>
        /// Returns value from specified column of data row
        /// </summary>
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