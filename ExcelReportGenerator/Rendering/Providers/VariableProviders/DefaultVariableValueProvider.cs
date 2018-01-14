using ExcelReportGenerator.Exceptions;
using ExcelReportGenerator.Helpers;
using System;
using System.Reflection;

namespace ExcelReportGenerator.Rendering.Providers.VariableProviders
{
    public class DefaultVariableValueProvider : IVariableValueProvider
    {
        private readonly IReflectionHelper _reflectionHelper;

        public DefaultVariableValueProvider() : this(new ReflectionHelper())
        {
        }

        internal DefaultVariableValueProvider(IReflectionHelper reflectionHelper)
        {
            _reflectionHelper = reflectionHelper;
        }

        public DateTime RenderDate { get; internal set; }

        public string SheetName { get; internal set; }

        public int SheetNumber { get; internal set; }

        public virtual object GetVariable(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(name));
            }

            PropertyInfo prop = _reflectionHelper.TryGetProperty(GetType(), name);
            if (prop == null)
            {
                throw new InvalidVariableException($"Cannot find variable with name \"{name}\" in class {GetType().Name} and all its parents. Variable must be public instance property.");
            }

            return prop.GetValue(this);
        }
    }
}