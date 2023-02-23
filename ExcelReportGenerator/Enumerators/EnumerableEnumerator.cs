using ExcelReportGenerator.Helpers;
using System.Collections;

namespace ExcelReportGenerator.Enumerators;

internal class EnumerableEnumerator : ICustomEnumerator
{
    private readonly IEnumerable _enumerable;
    private IEnumerator _enumerator;

    private int? _rowCount;

    public EnumerableEnumerator(IEnumerable enumerable)
    {
        _enumerable = enumerable ?? throw new ArgumentNullException(nameof(enumerable), ArgumentHelper.NullParamMessage);
        _enumerator = enumerable.GetEnumerator();
    }

    public bool MoveNext() => _enumerator.MoveNext();

    public void Reset()
    {
        try
        {
            _enumerator.Reset();
        }
        catch (Exception e) when (e is NotSupportedException || e is NotImplementedException)
        {
            _enumerator = _enumerable.GetEnumerator();
        }
    }

    public object Current => _enumerator.Current;

    public int RowCount
    {
        get
        {
            if (!_rowCount.HasValue)
            {
                if (_enumerable is ICollection collection)
                {
                    _rowCount = collection.Count;
                }
                else
                {
                    Type enumerableType = _enumerable.GetType();
                    Type genericCollectionInterface = TypeHelper.TryGetGenericCollectionInterface(enumerableType);
                    if (genericCollectionInterface != null)
                    {
                        _rowCount = (int)enumerableType.GetProperty(nameof(ICollection.Count)).GetValue(_enumerable, null);
                    }
                    else
                    {
                        int itemsCount = 0;
                        foreach (object _ in _enumerable)
                        {
                            itemsCount++;
                        }
                        _rowCount = itemsCount;
                    }
                }
            }

            return _rowCount.Value;
        }
    }
}