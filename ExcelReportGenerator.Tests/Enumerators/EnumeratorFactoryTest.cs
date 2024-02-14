using System.Collections;
using System.Data;
using ExcelReportGenerator.Enumerators;
using Assert = NUnit.Framework.Legacy.ClassicAssert;

namespace ExcelReportGenerator.Tests.Enumerators;

public class EnumeratorFactoryTest
{
    [Test]
    public void TestCreate()
    {
        Assert.IsNull(EnumeratorFactory.Create(null));

        Assert.IsInstanceOf<EnumerableEnumerator>(EnumeratorFactory.Create(new List<string>()));
        Assert.IsInstanceOf<EnumerableEnumerator>(EnumeratorFactory.Create(new int[0]));
        Assert.IsInstanceOf<EnumerableEnumerator>(EnumeratorFactory.Create(new Dictionary<string, object>()));
        Assert.IsInstanceOf<EnumerableEnumerator>(EnumeratorFactory.Create(new HashSet<string>()));
        Assert.IsInstanceOf<EnumerableEnumerator>(EnumeratorFactory.Create(new Hashtable()));
        Assert.IsInstanceOf<EnumerableEnumerator>(EnumeratorFactory.Create(new ArrayList()));

        var dataSet = new DataSet();
        dataSet.Tables.Add(new DataTable());

        Assert.IsInstanceOf<DataSetEnumerator>(EnumeratorFactory.Create(dataSet));
        Assert.IsInstanceOf<DataTableEnumerator>(EnumeratorFactory.Create(new DataTable()));

        var dataReader = Substitute.For<IDataReader>();
        Assert.IsInstanceOf<DataReaderEnumerator>(EnumeratorFactory.Create(dataReader));
    }
}