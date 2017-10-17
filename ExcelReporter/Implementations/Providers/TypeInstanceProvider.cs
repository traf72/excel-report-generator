using ExcelReporter.Interfaces.Providers;
using System;

namespace ExcelReporter.Implementations.Providers
{
    public class TypeInstanceProvider : ITypeInstanceProvider
    {
        private readonly ITypeProvider _typeProvider;

        public TypeInstanceProvider(ITypeProvider typeProvider)
        {
            if (typeProvider == null)
            {
                throw new ArgumentNullException(nameof(typeProvider), Constants.NullParamMessage);
            }
            _typeProvider = typeProvider;
        }

        public virtual object GetInstance(string typeTemplate)
        {
            Type type = _typeProvider.GetType(typeTemplate);
            return Activator.CreateInstance(type);
        }
    }
}