using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;

namespace ExcelReporter.Enumerators
{
    internal class DataReaderEnumerator : IEnumerator<IDataReader>
    {
        private const string ClosedDataReaderMessage = "DataReader has been closed";
        private const string FinishedEnumeratorMessage = "Enumerator has been finished";

        private readonly IDataReader _dataReader;
        private bool _isStarted;
        private bool _isFinished;

        public DataReaderEnumerator(IDataReader dataReader)
        {
            if (dataReader == null)
            {
                throw new ArgumentNullException(nameof(dataReader), Constants.NullParamMessage);
            }
            _dataReader = dataReader;
        }

        public IDataReader Current
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
                return _dataReader;
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
            return !_isFinished;
        }

        public void Reset()
        {
            throw new NotSupportedException($"{nameof(DataReaderEnumerator)} does not support reset method");
        }

        public void Dispose()
        {
            _dataReader.Close();
        }
    }
}