using System;

namespace ExcelReportGenerator.Rendering.Providers
{
    /// <summary>
    /// Provides instance of specified type
    /// </summary>
    public interface IInstanceProvider
    {
        object GetInstance(Type type);

        T GetInstance<T>();
    }
}