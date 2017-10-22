using ExcelReporter.Exceptions;
using ExcelReporter.Interfaces.Providers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelReporter.Attributes;

namespace ExcelReporter.Implementations.Providers
{
    public class ReflectionParameterProvider : IParameterProvider
    {
        protected readonly object ParamsContext;
        private List<MemberInfo> _typeParameters;

        public ReflectionParameterProvider(object paramsContext)
        {
            if (paramsContext == null)
            {
                throw new ArgumentNullException(nameof(paramsContext), Constants.NullParamMessage);
            }
            ParamsContext = paramsContext;
        }

        public object GetParameterValue(string paramName)
        {
            if (string.IsNullOrWhiteSpace(paramName))
            {
                throw new ArgumentException(Constants.EmptyStringParamMessage, nameof(paramName));
            }

            paramName = paramName.Trim();
            MemberInfo paramMember = AllTypeParameters.SingleOrDefault(p => p.Name == paramName);
            if (paramMember == null)
            {
                throw new ParameterNotFoundException($"Cannot find public instance property or field \"{paramName}\" with attribute \"{nameof(Parameter)}\" in type \"{ParamsContext.GetType().Name}\" and all its parents");
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

                _typeParameters = new List<MemberInfo>();
                Type paramsContextType = ParamsContext.GetType();
                const BindingFlags flags = BindingFlags.Instance | BindingFlags.Public;
                Func<MemberInfo, bool> whereClause = m => Attribute.IsDefined(m, typeof(Parameter));
                _typeParameters.AddRange(paramsContextType.GetProperties(flags).Where(whereClause));
                _typeParameters.AddRange(paramsContextType.GetFields(flags).Where(whereClause));
                return _typeParameters;
            }
        }
    }
}