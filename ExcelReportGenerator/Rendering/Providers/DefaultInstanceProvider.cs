using ExcelReportGenerator.Attributes;
using System;
using System.Collections.Generic;

namespace ExcelReportGenerator.Rendering.Providers
{
    /// <summary>
    /// Reflection type instance provider
    /// </summary>
    [LicenceKeyPart(L = true)]
    public class DefaultInstanceProvider : IInstanceProvider
    {
        private readonly IDictionary<Type, object> _instanceCache = new Dictionary<Type, object>();

        /// <param name="defaultInstance">Instance which will be returned if the type is not specified explicitly</param>
        public DefaultInstanceProvider(object defaultInstance = null)
        {
            DefaultInstance = defaultInstance;
            if (DefaultInstance != null)
            {
                _instanceCache[DefaultInstance.GetType()] = DefaultInstance;
            }
        }

        protected object DefaultInstance { get; }

        /// <summary>
        /// Provides instance of specified type as singleton. Type must have a default constructor.
        /// </summary>
        public virtual object GetInstance(Type type)
        {
            if (type == null)
            {
                return DefaultInstance ?? throw new InvalidOperationException("Type is not specified but defaultInstance is null");
            }

            if (_instanceCache.TryGetValue(type, out object instance))
            {
                return instance;
            }

            instance = Activator.CreateInstance(type);
            _instanceCache[type] = instance;
            return instance;
        }

        /// <summary>
        /// Provides instance of specified type as singleton. Type must have a default constructor.
        /// </summary>
        public virtual T GetInstance<T>()
        {
            return (T)GetInstance(typeof(T));
        }
    }
}