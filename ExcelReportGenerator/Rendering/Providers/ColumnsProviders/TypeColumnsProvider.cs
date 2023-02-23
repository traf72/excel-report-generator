using ExcelReportGenerator.Attributes;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Extensions;
using System.Reflection;

namespace ExcelReportGenerator.Rendering.Providers.ColumnsProviders;

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
            var columnAttr = Extensions.CustomAttributeExtensions.GetCustomAttribute<ExcelColumnAttribute>(probableColumnMember);
            if (columnAttr != null)
            {
                var excelColumn = new ExcelDynamicColumn(probableColumnMember.Name, memberType, columnAttr.Caption)
                {
                    Width = columnAttr.Width > 0 ? columnAttr.Width : (double?)null,
                    AdjustToContent = columnAttr.AdjustToContent,
                    Order = columnAttr.Order,
                };
                SetAggregationFunction(columnAttr, excelColumn);
                SetDisplayFormat(columnAttr, excelColumn);

                result.Add(excelColumn);
            }
            else if (memberType.IsExtendedPrimitive() || memberType.IsEnum)
            {
                result.Add(new ExcelDynamicColumn(probableColumnMember.Name, memberType));
            }
        }

        return result.OrderBy(c => c.Order).ToList();
    }

    private void SetAggregationFunction(ExcelColumnAttribute columnAttr, ExcelDynamicColumn column)
    {
        if (columnAttr.NoAggregate)
        {
            column.AggregateFunction = AggregateFunction.NoAggregation;
        }
        else if (columnAttr.AggregateFunction != AggregateFunction.NoAggregation)
        {
            column.AggregateFunction = columnAttr.AggregateFunction;
        }
    }

    private void SetDisplayFormat(ExcelColumnAttribute columnAttr, ExcelDynamicColumn column)
    {
        if (columnAttr.IgnoreDisplayFormat)
        {
            column.DisplayFormat = null;
        }
        else if (!string.IsNullOrWhiteSpace(columnAttr.DisplayFormat))
        {
            column.DisplayFormat = columnAttr.DisplayFormat;
        }
    }

    IList<ExcelDynamicColumn> IColumnsProvider.GetColumnsList(object type)
    {
        return GetColumnsList((Type)type);
    }
}