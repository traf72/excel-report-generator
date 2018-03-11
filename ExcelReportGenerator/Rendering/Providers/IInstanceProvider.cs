using System;
using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Rendering.Providers
{
    /// <summary>
    /// Provides instance of specified type
    /// </summary>
    [LicenceKeyPart]
    public interface IInstanceProvider
    {
        object GetInstance(Type type);

        T GetInstance<T>();
    }
}