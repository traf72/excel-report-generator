using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelReporter.Attributes;
using ExcelReporter.Exceptions;
using ExcelReporter.Helpers;

namespace ExcelReporter.Rendering.Providers
{
    /// <summary>
    /// Provides parameters values from instance members via reflection
    /// </summary>
    public class ReflectionParameterProvider : IParameterProvider
    {
        protected readonly object ParamsContext;
        private List<MemberInfo> _typeParameters;

        /// <param name="paramsContext">Object where parameters will be searched</param>
        public ReflectionParameterProvider(object paramsContext)
        {
            ParamsContext = paramsContext ?? throw new ArgumentNullException(nameof(paramsContext), ArgumentHelper.NullParamMessage);
        }

        public virtual object GetParameterValue(string paramName)
        {
            if (string.IsNullOrWhiteSpace(paramName))
            {
                throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(paramName));
            }

            paramName = paramName.Trim();
            MemberInfo paramMember = AllTypeParameters.SingleOrDefault(p => p.Name == paramName);
            if (paramMember == null)
            {
                throw new ParameterNotFoundException($"Cannot find public instance property or field \"{paramName}\" with attribute \"{nameof(ParameterAttribute)}\" in type \"{ParamsContext.GetType().Name}\" and all its parents");
            }

            return paramMember is PropertyInfo
                ? ((PropertyInfo)paramMember).GetValue(ParamsContext) : ((FieldInfo)paramMember).GetValue(ParamsContext);
        }

        private IEnumerable<MemberInfo> AllTypeParameters
        {
            get
            {
                if (_typeParameters != null)
                {
                    return _typeParameters;
                }

                bool GetOnlyParams(MemberInfo member) => Attribute.IsDefined(member, typeof(ParameterAttribute));

                _typeParameters = new List<MemberInfo>();
                Type paramsContextType = ParamsContext.GetType();
                const BindingFlags flags = BindingFlags.Instance | BindingFlags.Public;
                _typeParameters.AddRange(paramsContextType.GetProperties(flags).Where(GetOnlyParams));
                _typeParameters.AddRange(paramsContextType.GetFields(flags).Where(GetOnlyParams));
                return _typeParameters;
            }
        }
    }
}