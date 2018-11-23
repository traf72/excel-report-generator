using System;

namespace ExcelReportGenerator.Rendering.Providers
{
    /// <summary>
    /// Provides instances for templates
    /// </summary>
    public interface IInstanceProvider
    {
        /// <summary>
        /// Provides instance of specified <paramref name="type" />
        /// </summary>
        object GetInstance(Type type);

        /// <summary>
        /// Provides instance of specified <typeparamref name="T" />
        /// </summary>
        T GetInstance<T>();
    }
}