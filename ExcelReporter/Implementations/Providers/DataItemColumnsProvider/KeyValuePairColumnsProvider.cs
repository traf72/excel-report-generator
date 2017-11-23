using ExcelReporter.Interfaces.Providers.DataItemColumnsProvider;
using System.Collections.Generic;

namespace ExcelReporter.Implementations.Providers.DataItemColumnsProvider
{
    internal class KeyValuePairColumnsProvider : IDataItemColumnsProvider
    {
        public IList<ExcelDynamicColumn> GetColumnsList(object data)
        {
            return new List<ExcelDynamicColumn>
            {
                new ExcelDynamicColumn("Key"),
                new ExcelDynamicColumn("Value"),
            };
        }
    }
}