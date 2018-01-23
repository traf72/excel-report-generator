using System.Data;

namespace ExcelReportGenerator.Enumerators
{
    internal static class EnumeratorFactoryNew
    {
        public static ICustomEnumerator<DataRow> Create(object instance)
        {
            return new DataTableEnumerator((DataTable)instance);
        }
    }
}