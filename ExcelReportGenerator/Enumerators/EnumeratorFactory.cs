using System.Collections;
using System.Data;

namespace ExcelReportGenerator.Enumerators;

internal static class EnumeratorFactory
{
    public static ICustomEnumerator Create(object instance)
    {
        return instance switch
        {
            null => null,
            IDataReader dr => new DataReaderEnumerator(dr),
            DataTable dt => new DataTableEnumerator(dt),
            DataSet ds => new DataSetEnumerator(ds),
            IEnumerable e => new EnumerableEnumerator(e),
            _ => new EnumerableEnumerator(new[] {instance})
        };
    }
}