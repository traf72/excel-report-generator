using System.Reflection;

namespace ExcelReportGenerator.Helpers;

internal interface IReflectionHelper
{
    object GetValueOfPropertiesChain(string propertiesChain, object instance, BindingFlags flags = BindingFlags.Instance | BindingFlags.Public);

    PropertyInfo GetProperty(Type type, string propertyName, BindingFlags flags = BindingFlags.Instance | BindingFlags.Public);

    PropertyInfo TryGetProperty(Type type, string propertyName, BindingFlags flags = BindingFlags.Instance | BindingFlags.Public);

    FieldInfo GetField(Type type, string fieldName, BindingFlags flags = BindingFlags.Instance | BindingFlags.Public);

    FieldInfo TryGetField(Type type, string fieldName, BindingFlags flags = BindingFlags.Instance | BindingFlags.Public);

    object GetNullValueAttributeValue(MemberInfo member);
}