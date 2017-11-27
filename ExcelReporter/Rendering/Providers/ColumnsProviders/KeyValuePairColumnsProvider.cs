using System.Collections.Generic;

namespace ExcelReporter.Rendering.Providers.ColumnsProviders
{
    /// <summary>
    /// Provides columns info from KeyValuePair
    /// </summary>
    internal class KeyValuePairColumnsProvider : IColumnsProvider
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