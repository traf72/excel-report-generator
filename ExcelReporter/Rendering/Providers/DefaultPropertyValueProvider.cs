using ExcelReporter.Exceptions;
using ExcelReporter.Helpers;
using ExcelReporter.Rendering.TemplateParts;
using System;
using System.Linq;
using System.Reflection;

namespace ExcelReporter.Rendering.Providers
{
    /// <summary>
    /// Provides public properties or fields values via reflection
    /// </summary>
    public class DefaultPropertyValueProvider : IPropertyValueProvider
    {
        private const char PropertiesSeparator = '.';

        private readonly IReflectionHelper _reflectionHelper;

        /// <param name="typeProvider">Type provider which will be used for type search</param>
        /// <param name="instanceProvider">Instance provider which will be used to get instance of specified type</param>
        public DefaultPropertyValueProvider(ITypeProvider typeProvider, IInstanceProvider instanceProvider) : this(typeProvider, instanceProvider, new ReflectionHelper())
        {
        }

        internal DefaultPropertyValueProvider(ITypeProvider typeProvider, IInstanceProvider instanceProvider, IReflectionHelper reflectionHelper)
        {
            TypeProvider = typeProvider ?? throw new ArgumentNullException(nameof(typeProvider), ArgumentHelper.NullParamMessage);
            InstanceProvider = instanceProvider ?? throw new ArgumentNullException(nameof(instanceProvider), ArgumentHelper.NullParamMessage);
            _reflectionHelper = reflectionHelper;
        }

        protected ITypeProvider TypeProvider { get; }

        protected IInstanceProvider InstanceProvider { get; }

        public virtual object GetValue(string propertyTemplate)
        {
            if (string.IsNullOrWhiteSpace(propertyTemplate))
            {
                throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(propertyTemplate));
            }

            propertyTemplate = propertyTemplate.Trim();
            MemberTemplateParts templateParts = ParseTemplate(propertyTemplate);
            if (string.IsNullOrWhiteSpace(templateParts.MemberName))
            {
                throw new IncorrectTemplateException(string.Format(Constants.IncorrectTemplateMessage, propertyTemplate));
            }

            Type type = TypeProvider.GetType(templateParts.TypeName);
            string[] props = templateParts.MemberName.Split(PropertiesSeparator);
            object firstPropValue = GetFirstPropertyOrFieldValue(props[0], type);
            return props.Length == 1
                ? firstPropValue
                : _reflectionHelper.GetValueOfPropertiesChain(string.Join(PropertiesSeparator.ToString(), props.Skip(1)), firstPropValue);
        }

        private object GetFirstPropertyOrFieldValue(string propOrFieldName, Type type)
        {
            var flags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.FlattenHierarchy;
            object instance = null;
            PropertyInfo prop = _reflectionHelper.TryGetProperty(type, propOrFieldName, flags);
            if (prop != null)
            {
                MethodInfo getMethod = prop.GetGetMethod();
                if (getMethod == null)
                {
                    throw new InvalidOperationException($"Property \"{propOrFieldName}\" of type \"{type.Name}\" has no public getter");
                }
                if (!getMethod.IsStatic)
                {
                    instance = InstanceProvider.GetInstance(type);
                }
                return prop.GetValue(instance);
            }

            FieldInfo field = _reflectionHelper.TryGetField(type, propOrFieldName, flags);
            if (field != null)
            {
                if (!field.IsStatic)
                {
                    instance = InstanceProvider.GetInstance(type);
                }
                return field.GetValue(instance);
            }

            throw new MemberNotFoundException($"Cannot find property or field \"{propOrFieldName}\" in class \"{type.Name}\" and all its parents. BindingFlags = {flags}");
        }

        protected virtual MemberTemplateParts ParseTemplate(string template)
        {
            int typeSeparatorIndex = template.LastIndexOf(':');
            return typeSeparatorIndex == -1
                ? new MemberTemplateParts(null, template.Trim())
                : new MemberTemplateParts(template.Substring(0, typeSeparatorIndex).Trim(), template.Substring(typeSeparatorIndex + 1).Trim());
        }
    }
}