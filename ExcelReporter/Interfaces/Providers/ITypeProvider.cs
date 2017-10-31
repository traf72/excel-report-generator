using System;

namespace ExcelReporter.Interfaces.Providers
{
    public interface ITypeProvider
    {
        /// <summary>
        /// Provides type based on template
        /// </summary>
        Type GetType(string typeTemplate);
    }
}