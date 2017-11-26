using ExcelReporter.Interfaces.Providers.DataItemColumnsProvider;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ExcelReporter.Implementations.Providers.DataItemColumnsProvider
{
    /// <summary>
    /// Provides columns info from IDataReader
    /// </summary>
    internal class DataReaderColumnsProvider : IGenericDataItemColumnsProvider<IDataReader>
    {
        public IList<ExcelDynamicColumn> GetColumnsList(IDataReader reader)
        {
            DataTable schemaTable = reader?.GetSchemaTable();
            if (schemaTable == null)
            {
                return new List<ExcelDynamicColumn>();
            }

            return schemaTable.Rows
                .Cast<DataRow>()
                .Select(r => new ExcelDynamicColumn((string)r["ColumnName"]))
                .ToList();
        }

        IList<ExcelDynamicColumn> IDataItemColumnsProvider.GetColumnsList(object reader)
        {
            return GetColumnsList((IDataReader)reader);
        }
    }
}