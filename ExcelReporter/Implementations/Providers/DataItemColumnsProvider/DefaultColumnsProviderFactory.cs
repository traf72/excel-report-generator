using ExcelReporter.Interfaces.Providers.DataItemColumnsProvider;
using System.Collections.Generic;
using System.Data;

namespace ExcelReporter.Implementations.Providers.DataItemColumnsProvider
{
    internal class DefaultColumnsProviderFactory : IDataItemColumnsProviderFactory
    {
        public virtual IDataItemColumnsProvider Create(object data)
        {
            switch (data)
            {
                case null:
                    return null;

                case IDataReader _:
                    return new DataReaderColumnsProvider();

                case DataTable _:
                    return new DataTableColumnsProvider();

                case DataSet _:
                    return new DataSetColumnsProvider(new DataTableColumnsProvider());

                case KeyValuePair<object, object> _:
                case IEnumerable<KeyValuePair<object, object>> _:
                    return new KeyValuePairColumnsProvider();

                case IEnumerable<IDictionary<string, object>> _:
                    return new DictionaryColumnsProvider();

                case IEnumerable<object> _:
                    return new EnumerableColumnsProvider(new TypeColumnsProvider());
            }

            return new ObjectColumnsProvider(new TypeColumnsProvider());
        }
    }
}