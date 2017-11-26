﻿using System.Collections.Generic;

namespace ExcelReporter.Rendering.Providers.DataItemColumnsProviders
{
    /// <summary>
    /// Provides columns info from KeyValuePair
    /// </summary>
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