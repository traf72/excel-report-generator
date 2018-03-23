using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ExcelReportGenerator.Rendering.Providers.ColumnsProviders
{
    // Provides columns info from DataTable
    internal class DataTableColumnsProvider : IGenericColumnsProvider<DataTable>
    {
        public IList<ExcelDynamicColumn> GetColumnsList(DataTable dataTable)
        {
            if (dataTable == null)
            {
                return new List<ExcelDynamicColumn>();
            }

            return dataTable.Columns
                .Cast<DataColumn>()
                .Select(column => new ExcelDynamicColumn(column.ColumnName, column.DataType, column.Caption))
                .ToList();
        }

        IList<ExcelDynamicColumn> IColumnsProvider.GetColumnsList(object dataTable)
        {
            return GetColumnsList((DataTable)dataTable);
        }
    }
}