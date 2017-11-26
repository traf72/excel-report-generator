using System;
using System.Data;
using ExcelReporter.Exceptions;
using ExcelReporter.Helpers;

namespace ExcelReporter.Rendering.Providers.DataItemValueProviders
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
                throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(columnName));
            }

            _dataRow = dataRow ?? throw new ArgumentNullException(nameof(dataRow), ArgumentHelper.NullParamMessage);
            _columnName = columnName.Trim();
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