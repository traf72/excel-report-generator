using System.Collections;
using System.Data;

namespace ExcelReportGenerator.Enumerators
{
    internal static class EnumeratorFactory
    {
        public static IEnumerator Create(object instance)
        {
            switch (instance)
            {
                case null:
                    return null;
                case IDataReader dr:
                    return new DataReaderEnumerator(dr);
                case DataTable dt:
                    return dt.AsEnumerable().GetEnumerator();
                case DataSet ds:
                    return new DataSetEnumerator(ds);
                case IEnumerable e:
                    return e.GetEnumerator();
            }

            return new[] { instance }.GetEnumerator();
        }
    }
}