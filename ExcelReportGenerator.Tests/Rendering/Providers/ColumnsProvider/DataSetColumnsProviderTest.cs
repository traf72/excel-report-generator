using System.Data;
using ExcelReportGenerator.Rendering.Providers.ColumnsProviders;
using ExcelReportGenerator.Tests.CustomAsserts;
using Assert = NUnit.Framework.Legacy.ClassicAssert;

namespace ExcelReportGenerator.Tests.Rendering.Providers.ColumnsProvider;

public class DataSetColumnsProviderTest
{
    [Test]
    public void TestGetColumnsList()
    {
        var dataTableColumnsProvider = Substitute.For<IGenericColumnsProvider<DataTable>>();

        var dataTable1 = new DataTable("Table1");
        var dataTable2 = new DataTable("Table2");
        var dataSet = new DataSet
        {
            Tables = {dataTable1, dataTable2}
        };

        IColumnsProvider columnsProvider = new DataSetColumnsProvider(dataTableColumnsProvider);
        columnsProvider.GetColumnsList(dataSet);
        dataTableColumnsProvider.Received(1).GetColumnsList(dataTable1);

        dataTableColumnsProvider.ClearReceivedCalls();

        columnsProvider = new DataSetColumnsProvider(dataTableColumnsProvider, "Table2");
        columnsProvider.GetColumnsList(dataSet);
        dataTableColumnsProvider.Received(1).GetColumnsList(dataTable2);

        dataTableColumnsProvider.ClearReceivedCalls();

        columnsProvider = new DataSetColumnsProvider(dataTableColumnsProvider, "BadTable");
        Assert.AreEqual(0, columnsProvider.GetColumnsList(dataSet).Count);
        dataTableColumnsProvider.DidNotReceiveWithAnyArgs().GetColumnsList(Arg.Any<DataTable>());
    }

    [Test]
    public void TestGetColumnsListIfDataSetIsNullOrEmpty()
    {
        IColumnsProvider columnsProvider = new DataSetColumnsProvider(new DataTableColumnsProvider());
        Assert.AreEqual(0, columnsProvider.GetColumnsList(null).Count);
        Assert.AreEqual(0, columnsProvider.GetColumnsList(new DataSet()).Count);
    }

    [Test]
    public void TestGetColumnsListIfDataTableColumnsProviderIsNull()
    {
        ExceptionAssert.Throws<ArgumentNullException>(() => new DataSetColumnsProvider(null));
    }
}