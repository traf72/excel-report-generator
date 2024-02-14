using System.Collections;
using System.Collections.ObjectModel;
using ExcelReportGenerator.Enumerators;
using Assert = NUnit.Framework.Legacy.ClassicAssert;

namespace ExcelReportGenerator.Tests.Enumerators;

public class EnumerableEnumeratorTest
{
    [Test]
    public void TestEnumerator()
    {
        int[] array = {1, 2, 3};
        ICustomEnumerator enumerator = new EnumerableEnumerator(array);
        CheckEnumerator(enumerator, 3);

        IList<string> list = new List<string> {"One", "Two", "Three", "Four"};
        CheckEnumerator(new EnumerableEnumerator(list), 4);

        IEnumerable<Guid> collection = new Collection<Guid> {Guid.NewGuid(), Guid.NewGuid()};
        CheckEnumerator(new EnumerableEnumerator(collection), 2);

        IDictionary<string, object> dict = new Dictionary<string, object>
            {["One"] = 1, ["Two"] = "Two", ["Three"] = Guid.NewGuid()};
        CheckEnumerator(new EnumerableEnumerator(dict), 3);

        ISet<string> set = new HashSet<string> {"One", "Two"};
        CheckEnumerator(new EnumerableEnumerator(set), 2);

        IEnumerable<int> customEnumerable = new CustomEnumerable<int>(new[] {1, 2, 3, 4, 5});
        CheckEnumerator(new EnumerableEnumerator(customEnumerable), 5);
    }

    private void CheckEnumerator(ICustomEnumerator enumerator, int count)
    {
        Assert.AreEqual(count, enumerator.RowCount);

        IList<object> result = new List<object>();
        while (enumerator.MoveNext())
        {
            result.Add(enumerator.Current);
        }

        Assert.AreEqual(count, result.Count);
        Assert.AreEqual(count, enumerator.RowCount);

        enumerator.Reset();
        result.Clear();

        Assert.AreEqual(count, enumerator.RowCount);

        while (enumerator.MoveNext())
        {
            result.Add(enumerator.Current);
        }

        Assert.AreEqual(count, result.Count);
        Assert.AreEqual(count, enumerator.RowCount);
    }

    private class CustomEnumerable<T> : IEnumerable<T>
    {
        private readonly IEnumerable<T> _enumerable;

        public CustomEnumerable(IEnumerable<T> enumerable)
        {
            _enumerable = enumerable;
        }

        public IEnumerator<T> GetEnumerator()
        {
            return new CustomEnumerator<T>(_enumerable);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }

    private class CustomEnumerator<T> : IEnumerator<T>
    {
        private readonly IEnumerator<T> _enumerator;

        public CustomEnumerator(IEnumerable<T> enumerable)
        {
            _enumerator = enumerable.GetEnumerator();
        }

        public bool MoveNext()
        {
            return _enumerator.MoveNext();
        }

        public void Reset()
        {
            throw new NotSupportedException();
        }

        public T Current => _enumerator.Current;

        object IEnumerator.Current => Current;

        public void Dispose()
        {
            _enumerator.Dispose();
        }
    }
}