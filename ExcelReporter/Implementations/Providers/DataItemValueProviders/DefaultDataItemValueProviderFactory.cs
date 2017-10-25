using ExcelReporter.Interfaces.Providers.DataItemValueProviders;
using System.Data;

namespace ExcelReporter.Implementations.Providers.DataItemValueProviders
{
    internal class DefaultDataItemValueProviderFactory : IDataItemValueProviderFactory
    {
        public virtual IDataItemValueProvider Create(object data)
        {
            if (data == null)
            {
                return new ObjectPropertyValueProvider();
            }

            var dataRow = data as DataRow;
            if (dataRow != null)
            {
                return new DataRowValueProvider();
            }

            var dataReader = data as IDataReader;
            if (dataReader != null)
            {
                return new DataReaderValueProvider();
            }

            return new ObjectPropertyValueProvider();
        }
    }
}