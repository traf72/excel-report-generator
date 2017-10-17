using System;

namespace ExcelReporter.Interfaces.Providers
{
    public interface ITypeProvider
    {
        Type GetType(string typeTemplate);
    }
}