using System;

namespace ExcelReporter.Interfaces.Providers
{
    public interface IMethodContextProvider
    {
        object GetMethodContext(string typeName);
    }
}