using System;
using ExcelReportGenerator.Rendering.Providers;

namespace ExcelReportGenerator.Samples.Customizations
{
    public class CustomInstanceProvider : DefaultInstanceProvider
    {
        private readonly object _defaultInstance;

        public CustomInstanceProvider(object defaultInstance = null) : base(defaultInstance)
        {
            _defaultInstance = defaultInstance;
        }

        public override object GetInstance(Type type)
        {
            return type == null ? _defaultInstance : Ioc.Container.GetInstance(type);
        }

        public override T GetInstance<T>()
        {
            return (T)GetInstance(typeof(T));
        }
    }
}