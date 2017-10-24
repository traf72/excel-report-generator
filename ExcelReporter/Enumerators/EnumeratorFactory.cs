using System.Collections;
using System.Data;

namespace ExcelReporter.Enumerators
{
    internal static class EnumeratorFactory
    {
        public static IEnumerator Create(object instance)
        {
            if (instance == null)
            {
                return null;
            }

            var enumerable = instance as IEnumerable;
            if (enumerable != null)
            {
                return enumerable.GetEnumerator();
            }

            var dataTable = instance as DataTable;
            if (dataTable != null)
            {
                return dataTable.AsEnumerable().GetEnumerator();
            }

            var dataSet = instance as DataSet;
            if (dataSet != null)
            {
                return new DataSetEnumerator(dataSet);
            }

            return new[] { instance }.GetEnumerator();
        }
    }
}