using ExcelReportGenerator.Exceptions;
using ExcelReportGenerator.Helpers;
using System;
using System.Reflection;

namespace ExcelReportGenerator.Rendering.Providers.VariableProviders
{
    /// <summary>
    /// Provides values for system variable templates
    /// </summary>
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

        /// <summary>
        /// Report render date
        /// </summary>
        public DateTime RenderDate { get; set; }

        /// <summary>
        /// Name of worksheet that is currently render 
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// Number of worksheet that is currently render 
        /// </summary>
        public int SheetNumber { get; set; }

        /// <summary>
        /// Get variable value by its <paramref name="name"/>
        /// </summary>
        /// <exception cref="ArgumentException">Thrown when <paramref name="name" /> is null, empty string or whitespace</exception>
        /// <exception cref="InvalidVariableException"></exception>
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