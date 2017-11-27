using System.Collections.Generic;
using System.Linq;

namespace ExcelReporter.Rendering.Providers.ColumnsProviders
{
    /// <summary>
    /// Provides columns info from collection of IDictionary
    /// </summary>
    internal class DictionaryColumnsProvider : IGenericColumnsProvider<IEnumerable<IDictionary<string, object>>>
    {
        public IList<ExcelDynamicColumn> GetColumnsList(IEnumerable<IDictionary<string, object>> data)
        {
            IDictionary<string, object> firstElement = data?.FirstOrDefault();
            if (firstElement == null)
            {
                return new List<ExcelDynamicColumn>();
            }

            IList<ExcelDynamicColumn> result = new List<ExcelDynamicColumn>();
            foreach (KeyValuePair<string, object> pair in firstElement)
            {
                result.Add(new ExcelDynamicColumn(pair.Key));
            }

            return result;
        }

        IList<ExcelDynamicColumn> IColumnsProvider.GetColumnsList(object data)
        {
            return GetColumnsList((IEnumerable<IDictionary<string, object>>)data);
        }
    }
}