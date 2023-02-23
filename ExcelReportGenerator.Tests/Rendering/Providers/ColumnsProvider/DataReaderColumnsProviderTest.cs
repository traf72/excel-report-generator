using ExcelReportGenerator.Rendering.Providers.ColumnsProviders;

namespace ExcelReportGenerator.Tests.Rendering.Providers.ColumnsProvider;

public class DataReaderColumnsProviderTest
{
    [Test]
    public void TestGetColumnsList()
    {
        var dataReader = new DataProvider().GetAllCustomersDataReader();
        IColumnsProvider columnsProvider = new DataReaderColumnsProvider();
        var columns = columnsProvider.GetColumnsList(dataReader);

        Assert.AreEqual(6, columns.Count);

        Assert.AreEqual("Id", columns[0].Name);
        Assert.AreEqual("Id", columns[0].Caption);
        Assert.AreEqual(typeof(int), columns[0].DataType);
        Assert.IsNull(columns[0].Width);

        Assert.AreEqual("Name", columns[1].Name);
        Assert.AreEqual("Name", columns[1].Caption);
        Assert.AreEqual(typeof(string), columns[1].DataType);
        Assert.IsNull(columns[1].Width);

        Assert.AreEqual("IsVip", columns[2].Name);
        Assert.AreEqual("IsVip", columns[2].Caption);
        Assert.AreEqual(typeof(bool), columns[2].DataType);
        Assert.IsNull(columns[2].Width);

        Assert.AreEqual("Type", columns[3].Name);
        Assert.AreEqual("Type", columns[3].Caption);
        Assert.AreEqual(typeof(int), columns[3].DataType);
        Assert.IsNull(columns[3].Width);

        Assert.AreEqual("Description", columns[4].Name);
        Assert.AreEqual("Description", columns[4].Caption);
        Assert.AreEqual(typeof(string), columns[4].DataType);
        Assert.IsNull(columns[4].Width);

        Assert.AreEqual("Revenue", columns[5].Name);
        Assert.AreEqual("Revenue", columns[5].Caption);
        Assert.AreEqual(typeof(decimal), columns[5].DataType);
        Assert.IsNull(columns[5].Width);

        dataReader.Close();
    }

    [Test]
    public void TestGetColumnsListIfDataReaderIsEmpty()
    {
        var dataReader = new DataProvider().GetEmptyDataReader();
        IColumnsProvider columnsProvider = new DataReaderColumnsProvider();
        var columns = columnsProvider.GetColumnsList(dataReader);

        Assert.AreEqual(6, columns.Count);

        Assert.AreEqual("Id", columns[0].Name);
        Assert.AreEqual("Id", columns[0].Caption);
        Assert.AreEqual(typeof(int), columns[0].DataType);
        Assert.IsNull(columns[0].Width);

        Assert.AreEqual("Name", columns[1].Name);
        Assert.AreEqual("Name", columns[1].Caption);
        Assert.AreEqual(typeof(string), columns[1].DataType);
        Assert.IsNull(columns[1].Width);

        Assert.AreEqual("IsVip", columns[2].Name);
        Assert.AreEqual("IsVip", columns[2].Caption);
        Assert.AreEqual(typeof(bool), columns[2].DataType);
        Assert.IsNull(columns[2].Width);

        Assert.AreEqual("Type", columns[3].Name);
        Assert.AreEqual("Type", columns[3].Caption);
        Assert.AreEqual(typeof(int), columns[3].DataType);
        Assert.IsNull(columns[3].Width);

        Assert.AreEqual("Description", columns[4].Name);
        Assert.AreEqual("Description", columns[4].Caption);
        Assert.AreEqual(typeof(string), columns[4].DataType);
        Assert.IsNull(columns[4].Width);

        Assert.AreEqual("Revenue", columns[5].Name);
        Assert.AreEqual("Revenue", columns[5].Caption);
        Assert.AreEqual(typeof(decimal), columns[5].DataType);
        Assert.IsNull(columns[5].Width);

        dataReader.Close();
    }

    [Test]
    public void TestGetColumnsListIfDataReaderIsNull()
    {
        IColumnsProvider columnsProvider = new DataReaderColumnsProvider();
        Assert.AreEqual(0, columnsProvider.GetColumnsList(null).Count);
    }
}