﻿using ExcelReportGenerator.Attributes;
using ExcelReportGenerator.Exceptions;
using System.Dynamic;
using System.Reflection;

namespace ExcelReportGenerator.Helpers;

internal class ReflectionHelper : IReflectionHelper
{
    // Returns value of last property or field in the chain. Static properties and fields are not supported.
    /// <param name="propertiesChain">The chain where properties or fields names should be separated with ".". Example - Employee.Contact.Phone</param>
    public object GetValueOfPropertiesChain(string propertiesChain, object instance, BindingFlags flags = BindingFlags.Instance | BindingFlags.Public)
    {
        if (string.IsNullOrWhiteSpace(propertiesChain))
        {
            throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(propertiesChain));
        }
        if (flags.HasFlag(BindingFlags.Static))
        {
            throw new InvalidOperationException("BindingFlags.Static is specified but static properties and fields are not supported");
        }

        var queue = new Queue<string>(propertiesChain.Trim().Split('.'));
        int propsCount = queue.Count;
        int currentPropNumber = 0;
        while (queue.Count > 0)
        {
            currentPropNumber++;
            string propOrFieldName = queue.Dequeue();
            if (instance == null)
            {
                throw new NullReferenceException($"Cannot get property or field \"{propOrFieldName}\" because instance is null");
            }

            if (instance is ExpandoObject expando)
            {
                var dict = (IDictionary<string, object>)expando;
                if (dict.TryGetValue(propOrFieldName, out object value))
                {
                    instance = value;
                }
                else
                {
                    throw new MemberNotFoundException($"Cannot find property \"{propOrFieldName}\" in ExpandoObject");
                }
                continue;
            }

            PropertyInfo prop = TryGetProperty(instance.GetType(), propOrFieldName, flags);
            if (prop != null)
            {
                instance = prop.GetValue(instance, null);
                if (instance == null && currentPropNumber == propsCount)
                {
                    instance = GetNullValueAttributeValue(prop);
                }
                continue;
            }

            FieldInfo field = TryGetField(instance.GetType(), propOrFieldName, flags);
            if (field != null)
            {
                instance = field.GetValue(instance);
                if (instance == null && currentPropNumber == propsCount)
                {
                    instance = GetNullValueAttributeValue(field);
                }
                continue;
            }

            throw new MemberNotFoundException($"Cannot find property or field \"{propOrFieldName}\" in class \"{instance.GetType().Name}\" and all its parents. BindingFlags = {flags}");
        }

        return instance;
    }

    public PropertyInfo GetProperty(Type type, string propertyName, BindingFlags flags = BindingFlags.Instance | BindingFlags.Public)
    {
        if (type == null)
        {
            throw new ArgumentNullException(nameof(type), ArgumentHelper.NullParamMessage);
        }
        if (string.IsNullOrWhiteSpace(propertyName))
        {
            throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(propertyName));
        }

        PropertyInfo prop = type.GetProperty(propertyName, flags);
        if (prop == null)
        {
            throw new MemberNotFoundException($"Cannot find property \"{propertyName}\" in class \"{type.Name}\" and all its parents. BindingFlags = {flags}");
        }

        return prop;
    }

    public PropertyInfo TryGetProperty(Type type, string propertyName, BindingFlags flags = BindingFlags.Instance | BindingFlags.Public)
    {
        try
        {
            return GetProperty(type, propertyName, flags);
        }
        catch (MemberNotFoundException)
        {
            return null;
        }
    }

    public FieldInfo GetField(Type type, string fieldName, BindingFlags flags = BindingFlags.Instance | BindingFlags.Public)
    {
        if (type == null)
        {
            throw new ArgumentNullException(nameof(type), ArgumentHelper.NullParamMessage);
        }
        if (string.IsNullOrWhiteSpace(fieldName))
        {
            throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(fieldName));
        }

        FieldInfo field = type.GetField(fieldName, flags);
        if (field == null)
        {
            throw new MemberNotFoundException($"Cannot find field \"{fieldName}\" in class \"{type.Name}\" and all its parents. BindingFlags = {flags}");
        }

        return field;
    }

    public FieldInfo TryGetField(Type type, string fieldName, BindingFlags flags = BindingFlags.Instance | BindingFlags.Public)
    {
        try
        {
            return GetField(type, fieldName, flags);
        }
        catch (MemberNotFoundException)
        {
            return null;
        }
    }

    public object GetNullValueAttributeValue(MemberInfo member)
    {
        var nullValueAttr = Extensions.CustomAttributeExtensions.GetCustomAttribute<NullValueAttribute>(member);
        return nullValueAttr?.Value;
    }
}