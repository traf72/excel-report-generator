using System.Collections;
using System.Data;

namespace ExcelReporter.Enumerators
{
    internal static class EnumeratorFactory
    {
        public static IEnumerator Create(object instance)
        {
            switch (instance)
            {
                case null:
                    return null;
                case IEnumerable e:
                    return e.GetEnumerator();
                case DataTable dt:
                    return dt.AsEnumerable().GetEnumerator();
                case DataSet ds:
                    return new DataSetEnumerator(ds);
            }

            return new[] { instance }.GetEnumerator();
        }
    }
}