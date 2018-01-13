using ExcelReportGenerator.Exceptions;
using ExcelReportGenerator.Helpers;
using ExcelReportGenerator.Rendering.TemplateParts;
using System;
using System.Linq;
using System.Reflection;

namespace ExcelReportGenerator.Rendering.Providers
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
                throw new InvalidTemplateException(string.Format(Constants.InvalidTemplateMessage, propertyTemplate));
            }

            Type type = TypeProvider.GetType(templateParts.TypeName);
            string[] props = templateParts.MemberName.Split(PropertiesSeparator);
            MemberInfo member = GetFirstMember(props[0], type);
            object firstMemberValue = GetFirstMemberValue(member, type);

            if (props.Length == 1)
            {
                return firstMemberValue ?? _reflectionHelper.GetNullValueAttributeValue(member);
            }
            return _reflectionHelper.GetValueOfPropertiesChain(string.Join(PropertiesSeparator.ToString(), props.Skip(1)), firstMemberValue);
        }

        private MemberInfo GetFirstMember(string memberName, Type type)
        {
            var flags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.FlattenHierarchy;
            PropertyInfo prop = _reflectionHelper.TryGetProperty(type, memberName, flags);
            if (prop != null)
            {
                return prop;
            }

            FieldInfo field = _reflectionHelper.TryGetField(type, memberName, flags);
            if (field != null)
            {
                return field;
            }

            throw new MemberNotFoundException($"Cannot find property or field \"{memberName}\" in class \"{type.Name}\" and all its parents. BindingFlags = {flags}");
        }

        private object GetFirstMemberValue(MemberInfo member, Type type)
        {
            object instance = null;
            PropertyInfo prop = member as PropertyInfo;
            if (prop != null)
            {
                MethodInfo getMethod = prop.GetGetMethod();
                if (getMethod == null)
                {
                    throw new InvalidOperationException($"Property \"{prop.Name}\" of type \"{type.Name}\" has no public getter");
                }
                if (!getMethod.IsStatic)
                {
                    instance = InstanceProvider.GetInstance(type);
                }

                return prop.GetValue(instance);
            }

            FieldInfo field = member as FieldInfo;
            if (field != null)
            {
                if (!field.IsStatic)
                {
                    instance = InstanceProvider.GetInstance(type);
                }

                return field.GetValue(instance);
            }

            throw new ArgumentException($"Parameter must have the type of \"{nameof(PropertyInfo)}\" or {nameof(FieldInfo)}", nameof(member));
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