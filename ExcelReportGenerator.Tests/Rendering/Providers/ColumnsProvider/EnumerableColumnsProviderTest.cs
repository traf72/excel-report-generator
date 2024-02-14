using System.Collections;
using ExcelReportGenerator.Rendering.Providers.ColumnsProviders;
using ExcelReportGenerator.Tests.CustomAsserts;
using Assert = NUnit.Framework.Legacy.ClassicAssert;

namespace ExcelReportGenerator.Tests.Rendering.Providers.ColumnsProvider;

public class EnumerableColumnsProviderTest
{
    [Test]
    public void TestGetColumnsList()
    {
        var typeColumnsProvider = Substitute.For<IGenericColumnsProvider<Type>>();
        IColumnsProvider columnsProvider = new EnumerableColumnsProvider(typeColumnsProvider);

        columnsProvider.GetColumnsList(columnsProvider.GetColumnsList(new ArrayList
            {new TypeColumnsProviderTest.TestType(), "str"}));
        typeColumnsProvider.Received(1).GetColumnsList(typeof(TypeColumnsProviderTest.TestType));

        typeColumnsProvider.ClearReceivedCalls();

        columnsProvider.GetColumnsList(columnsProvider.GetColumnsList(new ArrayList
            {"str", new TypeColumnsProviderTest.TestType()}));
        typeColumnsProvider.Received(1).GetColumnsList(typeof(string));
    }

    [Test]
    public void TestGetColumnsListIfEnumerableIsNullOrEmpty()
    {
        IColumnsProvider columnsProvider = new EnumerableColumnsProvider(new TypeColumnsProvider());
        Assert.AreEqual(0, columnsProvider.GetColumnsList(null).Count);
        Assert.AreEqual(0, columnsProvider.GetColumnsList(new ArrayList()).Count);
    }

    [Test]
    public void TestGetColumnsListIfTypeColumnsProviderIsNull()
    {
        ExceptionAssert.Throws<ArgumentNullException>(() => new EnumerableColumnsProvider(null));
    }
}