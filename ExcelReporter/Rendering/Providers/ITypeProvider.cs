using System;

namespace ExcelReporter.Rendering.Providers
{
    public interface ITypeProvider
    {
        /// <summary>
        /// Provides type based on template
        /// </summary>
        Type GetType(string typeTemplate);
    }
}