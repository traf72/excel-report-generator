using ExcelReporter.Exceptions;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelReporter.Helpers
{
    internal class ReflectionHelper : IReflectionHelper
    {
        /// <summary>
        /// Returns value of last property or field in the chain. Static properties and fields are not supported.
        /// </summary>
        /// <param name="propertiesChain">Properties chain. Properties or fields names must separate with ".". Example - Employee.Contact.Phone</param>
        public object GetValueOfPropertiesChain(string propertiesChain, object instance, BindingFlags flags = BindingFlags.Instance | BindingFlags.Public)
        {
            if (string.IsNullOrWhiteSpace(propertiesChain))
            {
                throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(propertiesChain));
            }
            if ((flags & BindingFlags.Static) == BindingFlags.Static)
            {
                throw new InvalidOperationException("BindingFlags.Static is specified but static properties and fields are not supported");
            }

            Queue<string> queue = new Queue<string>(propertiesChain.Trim().Split('.'));
            while (queue.Count > 0)
            {
                string propOrFieldName = queue.Dequeue();
                if (instance == null)
                {
                    throw new InvalidOperationException($"Cannot get property or field \"{propOrFieldName}\" because instance is null");
                }

                PropertyInfo prop = TryGetProperty(instance.GetType(), propOrFieldName, flags);
                if (prop != null)
                {
                    instance = prop.GetValue(instance);
                    continue;
                }

                FieldInfo field = TryGetField(instance.GetType(), propOrFieldName, flags);
                if (field != null)
                {
                    instance = field.GetValue(instance);
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
    }
}