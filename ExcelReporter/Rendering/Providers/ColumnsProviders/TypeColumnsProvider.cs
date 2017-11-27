using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelReporter.Attributes;

namespace ExcelReporter.Rendering.Providers.ColumnsProviders
{
    /// <summary>
    /// Provides columns info from from Type
    /// </summary>
    internal class TypeColumnsProvider : IGenericColumnsProvider<Type>
    {
        public IList<ExcelDynamicColumn> GetColumnsList(Type type)
        {
            if (type == null)
            {
                return new List<ExcelDynamicColumn>();
            }

            BindingFlags flags = BindingFlags.Instance | BindingFlags.Public;
            MemberInfo[] excelColumns = type.GetProperties(flags)
                // Пока исключил поля, так как шаблоны пока парсятся только из свойств
                //.AsEnumerable<MemberInfo>()
                //.Concat(dataItemType.GetFields(flags))
                .Where(m => m.IsDefined(typeof(ExcelColumnAttribute), true)).ToArray();

            IList<ExcelDynamicColumn> result = new List<ExcelDynamicColumn>();
            foreach (MemberInfo columnMember in excelColumns)
            {
                var columnAttr = (ExcelColumnAttribute)columnMember.GetCustomAttribute(typeof(ExcelColumnAttribute), true);
                result.Add(new ExcelDynamicColumn(columnMember.Name, columnAttr.Caption) { Width = columnAttr.Width > 0 ? columnAttr.Width : (double?)null });
            }

            return result;
        }

        IList<ExcelDynamicColumn> IColumnsProvider.GetColumnsList(object type)
        {
            return GetColumnsList((Type)type);
        }
    }
}