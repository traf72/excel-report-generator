using ExcelReportGenerator.Rendering.Providers.ColumnsProviders;
using ExcelReportGenerator.Tests.CustomAsserts;
using Assert = NUnit.Framework.Legacy.ClassicAssert;

namespace ExcelReportGenerator.Tests.Rendering.Providers.ColumnsProvider;

public class ObjectColumnsProviderTest
{
    [Test]
    public void TestGetColumnsList()
    {
        var typeColumnsProvider = Substitute.For<IGenericColumnsProvider<Type>>();

        IColumnsProvider columnsProvider = new ObjectColumnsProvider(typeColumnsProvider);
        var testObject = new TypeColumnsProviderTest.TestType();
        columnsProvider.GetColumnsList(testObject);
        typeColumnsProvider.Received(1).GetColumnsList(testObject.GetType());

        typeColumnsProvider.ClearReceivedCalls();

        var str = "str";
        columnsProvider.GetColumnsList(str);
        typeColumnsProvider.Received(1).GetColumnsList(str.GetType());
    }

    [Test]
    public void TestGetColumnsListIfObjectIsNull()
    {
        IColumnsProvider columnsProvider = new ObjectColumnsProvider(new TypeColumnsProvider());
        Assert.AreEqual(0, columnsProvider.GetColumnsList(null).Count);
    }

    [Test]
    public void TestGetColumnsListIfTypeColumnsProviderIsNull()
    {
        ExceptionAssert.Throws<ArgumentNullException>(() => new ObjectColumnsProvider(null));
    }
}