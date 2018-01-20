using ExcelReportGenerator.Exceptions;
using ExcelReportGenerator.Helpers;
using System;
using System.Reflection;

namespace ExcelReportGenerator.Rendering.Providers.VariableProviders
{
    public class SystemVariableProvider
    {
        private readonly IReflectionHelper _reflectionHelper;

        public SystemVariableProvider() : this(new ReflectionHelper())
        {
        }

        internal SystemVariableProvider(IReflectionHelper reflectionHelper)
        {
            _reflectionHelper = reflectionHelper;
        }

        public DateTime RenderDate { get; set; }

        public string SheetName { get; set; }

        public int SheetNumber { get; set; }

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