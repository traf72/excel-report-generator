using ExcelReporter.Interfaces.Providers.DataItemColumnsProvider;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReporter.Implementations.Providers.DataItemColumnsProvider
{
    internal class DictionaryColumnsProvider : IGenericDataItemColumnsProvider<IEnumerable<IDictionary<string, object>>>
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

        IList<ExcelDynamicColumn> IDataItemColumnsProvider.GetColumnsList(object data)
        {
            return GetColumnsList((IEnumerable<IDictionary<string, object>>)data);
        }
    }
}