using System.Data;

namespace ExcelReportGenerator.Samples.Extensions
{
    public static class DataReaderExtensions
    {
        public static T GetValueSafe<T>(this IDataReader reader, int columnIndex)
        {
            return reader.IsDBNull(columnIndex) ? default(T) : (T)reader.GetValue(columnIndex);
        }
    }
}