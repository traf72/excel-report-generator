using ExcelReporter.Exceptions;
using ExcelReporter.Interfaces.Providers.DataItemValueProviders;
using System;
using System.Data;

namespace ExcelReporter.Implementations.Providers.DataItemValueProviders
{
    /// <summary>
    /// Provides values from data reader
    /// </summary>
    public class DataReaderValueProvider : IGenericDataItemValueProvider<IDataReader>
    {
        private string _columnName;
        private IDataReader _dataReader;

        /// <summary>
        /// Returns value from specified column of data reader
        /// </summary>
        public virtual object GetValue(string columnName, IDataReader dataReader)
        {
            if (string.IsNullOrWhiteSpace(columnName))
            {
                throw new ArgumentException(Constants.EmptyStringParamMessage, nameof(columnName));
            }
            if (dataReader == null)
            {
                throw new ArgumentNullException(nameof(dataReader), Constants.NullParamMessage);
            }
            if (dataReader.IsClosed)
            {
                throw new InvalidOperationException("DataReader is closed");
            }

            _columnName = columnName.Trim();
            _dataReader = dataReader;

            return dataReader.GetValue(GetColumnIndex());
        }

        private int GetColumnIndex()
        {
            int colIndex = -1;
            try
            {
                colIndex = _dataReader.GetOrdinal(_columnName);
            }
            catch (IndexOutOfRangeException e)
            {
                ThrowColumnNotFoundException(e);
            }

            if (colIndex == -1)
            {
                ThrowColumnNotFoundException();
            }
            return colIndex;
        }

        private void ThrowColumnNotFoundException(Exception innerException = null)
        {
            throw new ColumnNotFoundException($"DataReader does not contain column \"{_columnName}\"", innerException);
        }

        object IDataItemValueProvider.GetValue(string columnName, object dataReader)
        {
            return GetValue(columnName, (IDataReader)dataReader);
        }
    }
}