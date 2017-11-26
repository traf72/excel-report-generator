﻿using System;
using System.Data;
using ExcelReporter.Rendering.Providers.DataItemColumnsProviders;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;

namespace ExcelReporter.Tests.Rendering.Providers.DataItemColumnsProvider
{
    [TestClass]
    public class DataSetColumnsProviderTest
    {
        [TestMethod]
        public void TestGetColumnsList()
        {
            IGenericDataItemColumnsProvider<DataTable> dataTableColumsProvider = Substitute.For<IGenericDataItemColumnsProvider<DataTable>>();

            var dataTable1 = new DataTable("Table1");
            var dataTable2 = new DataTable("Table2");
            var dataSet = new DataSet
            {
                Tables = { dataTable1, dataTable2 }
            };

            IDataItemColumnsProvider columnsProvider = new DataSetColumnsProvider(dataTableColumsProvider);
            columnsProvider.GetColumnsList(dataSet);
            dataTableColumsProvider.Received(1).GetColumnsList(dataTable1);

            dataTableColumsProvider.ClearReceivedCalls();

            columnsProvider = new DataSetColumnsProvider(dataTableColumsProvider, "Table2");
            columnsProvider.GetColumnsList(dataSet);
            dataTableColumsProvider.Received(1).GetColumnsList(dataTable2);

            dataTableColumsProvider.ClearReceivedCalls();

            columnsProvider = new DataSetColumnsProvider(dataTableColumsProvider, "BadTable");
            Assert.AreEqual(0, columnsProvider.GetColumnsList(dataSet).Count);
            dataTableColumsProvider.DidNotReceiveWithAnyArgs().GetColumnsList(Arg.Any<DataTable>());
        }

        [TestMethod]
        public void TestGetColumnsListIfDataSetIsNullOrEmpty()
        {
            IDataItemColumnsProvider columnsProvider = new DataSetColumnsProvider(new DataTableColumnsProvider());
            Assert.AreEqual(0, columnsProvider.GetColumnsList(null).Count);
            Assert.AreEqual(0, columnsProvider.GetColumnsList(new DataSet()).Count);
        }

        [TestMethod]
        public void TestGetColumnsListIfDataTableColumnsProviderIsNull()
        {
            ExceptionAssert.Throws<ArgumentNullException>(() => new DataSetColumnsProvider(null));
        }
    }
}