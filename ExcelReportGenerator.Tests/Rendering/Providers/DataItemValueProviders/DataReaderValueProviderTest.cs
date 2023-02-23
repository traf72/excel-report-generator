using System.Data;
using ExcelReportGenerator.Exceptions;
using ExcelReportGenerator.Rendering.Providers.DataItemValueProviders;
using ExcelReportGenerator.Tests.CustomAsserts;
using NSubstitute.ExceptionExtensions;

namespace ExcelReportGenerator.Tests.Rendering.Providers.DataItemValueProviders;

public class DataReaderValueProviderTest
{
    [Test]
    public void TestGetValue()
    {
        IGenericDataItemValueProvider<IDataReader> provider = new DataReaderValueProvider();

        var dataReader = Substitute.For<IDataReader>();
        dataReader.GetOrdinal("Column1").Returns(0);
        dataReader.GetOrdinal("column1").Returns(0);
        dataReader.GetOrdinal("Column2").Returns(1);
        dataReader.GetOrdinal("BadColumn").Returns(-1);
        dataReader.GetOrdinal("BadColumn2").Throws(new IndexOutOfRangeException());

        dataReader.GetValue(0).Returns(5);
        dataReader.GetValue(1).Returns("Five");

        Assert.AreEqual(5, provider.GetValue("Column1", dataReader));
        Assert.AreEqual(5, provider.GetValue("column1", dataReader));
        Assert.AreEqual(5, provider.GetValue(" column1 ", dataReader));
        Assert.AreEqual("Five", provider.GetValue("Column2", dataReader));

        ExceptionAssert.Throws<ColumnNotFoundException>(() => provider.GetValue("BadColumn", dataReader),
            "DataReader does not contain column \"BadColumn\"");
        ExceptionAssert.Throws<ColumnNotFoundException>(() => provider.GetValue("BadColumn2", dataReader),
            "DataReader does not contain column \"BadColumn2\"");
        ExceptionAssert.Throws<ArgumentException>(() => provider.GetValue(null, dataReader));
        ExceptionAssert.Throws<ArgumentException>(() => provider.GetValue(string.Empty, dataReader));
        ExceptionAssert.Throws<ArgumentException>(() => provider.GetValue(" ", dataReader));
        ExceptionAssert.Throws<ArgumentNullException>(() => provider.GetValue("Column1", null));

        dataReader.IsClosed.Returns(true);
        ExceptionAssert.Throws<InvalidOperationException>(() => provider.GetValue("Column1", dataReader),
            "DataReader is closed");
    }

    [Test]
    public void TestGetValueWithRealSqlReader()
    {
        IGenericDataItemValueProvider<IDataReader> provider = new DataReaderValueProvider();
        var reader = new DataProvider().GetAllCustomersDataReader();

        reader.Read();
        Assert.AreEqual(1, provider.GetValue("Id", reader));
        Assert.AreEqual(1, provider.GetValue("id", reader));
        Assert.AreEqual("Customer 1", provider.GetValue("Name", reader));
        Assert.AreEqual(false, provider.GetValue("IsVip", reader));
        Assert.IsNull(provider.GetValue("Type", reader));

        reader.Read();
        Assert.AreEqual(2, provider.GetValue("Id", reader));
        Assert.AreEqual("Customer 2", provider.GetValue("Name", reader));
        Assert.AreEqual(true, provider.GetValue("IsVip", reader));
        Assert.AreEqual(1, provider.GetValue("Type", reader));

        reader.Read();
        Assert.AreEqual(3, provider.GetValue("Id", reader));
        Assert.AreEqual("Customer 3", provider.GetValue("Name", reader));
        Assert.IsNull(provider.GetValue("IsVip", reader));
        Assert.IsNull(provider.GetValue("Type", reader));

        reader.Close();
    }
}