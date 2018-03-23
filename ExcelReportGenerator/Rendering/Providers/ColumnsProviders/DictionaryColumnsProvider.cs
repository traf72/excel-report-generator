using System.Collections.Generic;
using System.Linq;

namespace ExcelReportGenerator.Rendering.Providers.ColumnsProviders
{
    // Provides columns info from collection of IDictionary
    internal class DictionaryColumnsProvider<TValue> : IGenericColumnsProvider<IEnumerable<IDictionary<string, TValue>>>
    {
        public IList<ExcelDynamicColumn> GetColumnsList(IEnumerable<IDictionary<string, TValue>> data)
        {
            IDictionary<string, TValue> firstElement = data?.FirstOrDefault();
            if (firstElement == null)
            {
                return new List<ExcelDynamicColumn>();
            }

            IList<ExcelDynamicColumn> result = new List<ExcelDynamicColumn>();
            foreach (KeyValuePair<string, TValue> pair in firstElement)
            {
                result.Add(new ExcelDynamicColumn(pair.Key, pair.Value?.GetType()));
            }

            return result;
        }

        IList<ExcelDynamicColumn> IColumnsProvider.GetColumnsList(object data)
        {
            return GetColumnsList((IEnumerable<IDictionary<string, TValue>>)data);
        }
    }
}