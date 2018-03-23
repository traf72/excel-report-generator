using System;
using System.Data;
using ExcelReportGenerator.Exceptions;
using ExcelReportGenerator.Helpers;

namespace ExcelReportGenerator.Rendering.Providers.DataItemValueProviders
{
    // Provides values from data row
    internal class DataRowValueProvider : IGenericDataItemValueProvider<DataRow>
    {
        private string _columnName;
        private DataRow _dataRow;

        // Returns value from specified column of data row
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