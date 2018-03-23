using System.Collections.Generic;

namespace ExcelReportGenerator.Rendering.Providers.ColumnsProviders
{
    internal interface IColumnsProvider
    {
        // Provides columns info based on data
        IList<ExcelDynamicColumn> GetColumnsList(object data);
    }
}