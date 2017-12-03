using ExcelReporter.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

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
            MemberInfo[] excelColumns = type.GetFields(flags)
                .AsEnumerable<MemberInfo>()
                .Concat(type.GetProperties(flags))
                .Where(m => m.IsDefined(typeof(ExcelColumnAttribute), true)).ToArray();

            IList<ExcelDynamicColumn> result = new List<ExcelDynamicColumn>();
            foreach (MemberInfo columnMember in excelColumns)
            {
                var columnAttr = (ExcelColumnAttribute)columnMember.GetCustomAttribute(typeof(ExcelColumnAttribute), true);
                Type columnType = columnMember is PropertyInfo p ? p.PropertyType : ((FieldInfo) columnMember).FieldType;
                result.Add(new ExcelDynamicColumn(columnMember.Name, columnType, columnAttr.Caption) { Width = columnAttr.Width > 0 ? columnAttr.Width : (double?)null });
            }

            return result;
        }

        IList<ExcelDynamicColumn> IColumnsProvider.GetColumnsList(object type)
        {
            return GetColumnsList((Type)type);
        }
    }
}