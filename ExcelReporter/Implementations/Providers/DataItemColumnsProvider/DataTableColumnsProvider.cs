using ExcelReporter.Interfaces.Providers.DataItemColumnsProvider;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ExcelReporter.Implementations.Providers.DataItemColumnsProvider
{
    /// <summary>
    /// Provides columns info from DataTable
    /// </summary>
    internal class DataTableColumnsProvider : IGenericDataItemColumnsProvider<DataTable>
    {
        public IList<ExcelDynamicColumn> GetColumnsList(DataTable dataTable)
        {
            if (dataTable == null)
            {
                return new List<ExcelDynamicColumn>();
            }

            return dataTable.Columns
                .Cast<DataColumn>()
                // TODO Определять DataType из column.DataType, а может и не нужно
                .Select(column => new ExcelDynamicColumn(column.ColumnName, column.Caption))
                .ToList();
        }

        IList<ExcelDynamicColumn> IDataItemColumnsProvider.GetColumnsList(object dataTable)
        {
            return GetColumnsList((DataTable)dataTable);
        }
    }
}