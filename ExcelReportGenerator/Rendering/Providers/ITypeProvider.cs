using System;

namespace ExcelReportGenerator.Rendering.Providers
{
    /// <summary>
    /// Provides types for templates
    /// </summary>
    public interface ITypeProvider
    {
        /// <summary>
        /// Provides type based on <paramref name="typeTemplate" />
        /// </summary>
        Type GetType(string typeTemplate);
    }
}