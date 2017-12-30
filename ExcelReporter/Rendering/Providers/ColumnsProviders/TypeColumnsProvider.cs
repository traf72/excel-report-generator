using ExcelReporter.Attributes;
using ExcelReporter.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelReporter.Enums;

namespace ExcelReporter.Rendering.Providers.ColumnsProviders
{
    /// <summary>
    /// Provides columns info from Type
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
            MemberInfo[] probableExcelColumns = type.GetFields(flags)
                .AsEnumerable<MemberInfo>()
                .Concat(type.GetProperties(flags))
                .Where(m => !m.IsDefined(typeof(NoExcelColumnAttribute), true)).ToArray();

            IList<ExcelDynamicColumn> result = new List<ExcelDynamicColumn>();
            foreach (MemberInfo probableColumnMember in probableExcelColumns)
            {
                Type memberType = probableColumnMember is PropertyInfo p ? p.PropertyType : ((FieldInfo)probableColumnMember).FieldType;
                var columnAttr = (ExcelColumnAttribute)probableColumnMember.GetCustomAttribute(typeof(ExcelColumnAttribute));
                if (columnAttr != null)
                {
                    var excelColumn =
                        new ExcelDynamicColumn(probableColumnMember.Name, memberType, columnAttr.Caption)
                        {
                            Width = columnAttr.Width > 0 ? columnAttr.Width : (double?)null,
                            AggregateFunction = columnAttr.NoAggregate ? AggregateFunction.NoAggregation : columnAttr.AggregateFunction,
                            DisplayFormat = columnAttr.IgnoreDisplayFormat ? null : columnAttr.DisplayFormat,
                            AdjustToContent = columnAttr.AdjustToContent,
                            Order = columnAttr.Order,
                        };
                    result.Add(excelColumn);
                }
                else if (memberType.IsExtendedPrimitive())
                {
                    result.Add(new ExcelDynamicColumn(probableColumnMember.Name, memberType));
                }
            }

            return result.OrderBy(c => c.Order).ToList();
        }

        IList<ExcelDynamicColumn> IColumnsProvider.GetColumnsList(object type)
        {
            return GetColumnsList((Type)type);
        }
    }
}