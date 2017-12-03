using System;

namespace ExcelReporter.Rendering.Providers
{
    /// <summary>
    /// Provides instance of specified type
    /// </summary>
    public interface IInstanceProvider
    {
        object GetInstance(Type type);
    }
}