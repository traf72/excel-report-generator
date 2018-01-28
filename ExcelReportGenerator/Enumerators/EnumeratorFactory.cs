using System.Collections;
using System.Data;

namespace ExcelReportGenerator.Enumerators
{
    internal static class EnumeratorFactory
    {
        public static ICustomEnumerator Create(object instance)
        {
            switch (instance)
            {
                case null:
                    return null;
                case IDataReader dr:
                    return new DataReaderEnumerator(dr);
                case DataTable dt:
                    return new DataTableEnumerator(dt);
                case DataSet ds:
                    return new DataSetEnumerator(ds);
                case IEnumerable e:
                    return new EnumerableEnumerator(e);
            }

            return new EnumerableEnumerator(new[] { instance });
        }
    }
}