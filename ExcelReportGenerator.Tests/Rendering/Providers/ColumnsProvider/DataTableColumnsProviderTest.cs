﻿using System.Data;
using ExcelReportGenerator.Rendering.Providers.ColumnsProviders;

namespace ExcelReportGenerator.Tests.Rendering.Providers.ColumnsProvider;

public class DataTableColumnsProviderTest
{
    [Test]
    public void TestGetColumnsList()
    {
        var dataTable = GetDataTable();
        dataTable.Rows.Add(1, "One", true);
        dataTable.Rows.Add(2, "Two", false);
        dataTable.Rows.Add(3, "Three", true);

        IColumnsProvider columnsProvider = new DataTableColumnsProvider();
        var columns = columnsProvider.GetColumnsList(dataTable);

        Assert.AreEqual(3, columns.Count);

        Assert.AreEqual("Column1", columns[0].Name);
        Assert.AreEqual("Column1", columns[0].Caption);
        Assert.AreEqual(typeof(string), columns[0].DataType);
        Assert.IsNull(columns[0].Width);

        Assert.AreEqual("Column2", columns[1].Name);
        Assert.AreEqual("Caption2", columns[1].Caption);
        Assert.AreEqual(typeof(string), columns[1].DataType);
        Assert.IsNull(columns[1].Width);

        Assert.AreEqual("Column3", columns[2].Name);
        Assert.AreEqual("Caption3", columns[2].Caption);
        Assert.AreEqual(typeof(bool), columns[2].DataType);
        Assert.IsNull(columns[2].Width);
    }

    [Test]
    public void TestGetColumnsListIfDataTableIsEmpty()
    {
        var dataTable = GetDataTable();

        IColumnsProvider columnsProvider = new DataTableColumnsProvider();
        var columns = columnsProvider.GetColumnsList(dataTable);

        Assert.AreEqual(3, columns.Count);

        Assert.AreEqual("Column1", columns[0].Name);
        Assert.AreEqual("Column1", columns[0].Caption);
        Assert.AreEqual(typeof(string), columns[0].DataType);
        Assert.IsNull(columns[0].Width);

        Assert.AreEqual("Column2", columns[1].Name);
        Assert.AreEqual("Caption2", columns[1].Caption);
        Assert.AreEqual(typeof(string), columns[1].DataType);
        Assert.IsNull(columns[1].Width);

        Assert.AreEqual("Column3", columns[2].Name);
        Assert.AreEqual("Caption3", columns[2].Caption);
        Assert.AreEqual(typeof(bool), columns[2].DataType);
        Assert.IsNull(columns[2].Width);
    }

    [Test]
    public void TestGetColumnsListIfDataTableIsNull()
    {
        IColumnsProvider columnsProvider = new DataTableColumnsProvider();
        Assert.AreEqual(0, columnsProvider.GetColumnsList(null).Count);
    }

    private DataTable GetDataTable()
    {
        var dataTable = new DataTable();
        dataTable.Columns.Add("Column1");
        dataTable.Columns.Add(new DataColumn("Column2", typeof(string)) {Caption = "Caption2"});
        dataTable.Columns.Add(new DataColumn("Column3", typeof(bool)) {Caption = "Caption3"});
        return dataTable;
    }
}