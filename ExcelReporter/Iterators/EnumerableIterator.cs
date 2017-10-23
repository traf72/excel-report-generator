using System;
using System.Collections.Generic;

namespace ExcelReporter.Iterators
{
    public class EnumerableIterator<T> : IIterator<T>
    {
        private readonly IEnumerator<T> _enumerableEnumerator;

        public EnumerableIterator(IEnumerable<T> enumerable)
        {
            if (enumerable == null)
            {
                throw new ArgumentNullException(nameof(enumerable), Constants.NullParamMessage);
            }
            _enumerableEnumerator = enumerable.GetEnumerator();
        }

        public T Next()
        {
            return _enumerableEnumerator.Current;
        }

        public bool HaxNext()
        {
            return _enumerableEnumerator.MoveNext();
        }

        public void Reset()
        {
            _enumerableEnumerator.Reset();
        }
    }
}