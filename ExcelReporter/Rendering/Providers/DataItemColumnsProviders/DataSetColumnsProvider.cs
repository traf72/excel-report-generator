﻿using System;
using System.Collections.Generic;
using System.Data;
using ExcelReporter.Helpers;

namespace ExcelReporter.Rendering.Providers.DataItemColumnsProviders
{
    /// <summary>
    /// Provides columns info from DataSet
    /// </summary>
    internal class DataSetColumnsProvider : IGenericDataItemColumnsProvider<DataSet>
    {
        private readonly IGenericDataItemColumnsProvider<DataTable> _dataTableColumnsProvider;
        private readonly string _tableName;

        public DataSetColumnsProvider(IGenericDataItemColumnsProvider<DataTable> dataTableColumnsProvider, string tableName = null)
        {
            _dataTableColumnsProvider = dataTableColumnsProvider ?? throw new ArgumentNullException(nameof(dataTableColumnsProvider), ArgumentHelper.NullParamMessage);
            _tableName = tableName;
        }

        public IList<ExcelDynamicColumn> GetColumnsList(DataSet dataSet)
        {
            if (dataSet == null || dataSet.Tables.Count == 0)
            {
                return new List<ExcelDynamicColumn>();
            }

            if (string.IsNullOrWhiteSpace(_tableName))
            {
                return _dataTableColumnsProvider.GetColumnsList(dataSet.Tables[0]);
            }

            DataTable table = dataSet.Tables[_tableName];
            return table == null ? new List<ExcelDynamicColumn>() : _dataTableColumnsProvider.GetColumnsList(table);
        }

        IList<ExcelDynamicColumn> IDataItemColumnsProvider.GetColumnsList(object dataSet)
        {
            return GetColumnsList((DataSet)dataSet);
        }
    }
}