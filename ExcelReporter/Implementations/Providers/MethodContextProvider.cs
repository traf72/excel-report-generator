using ExcelReporter.Interfaces.Providers;
using System;
using System.Linq;
using System.Reflection;
using ExcelReporter.Exceptions;
using ExcelReporter.Interfaces.Reports;

namespace ExcelReporter.Implementations.Providers
{
    public class MethodContextProvider : IMethodContextProvider
    {
        private readonly IReport _report;

        public MethodContextProvider(IReport report)
        {
            if (report == null)
            {
                throw new ArgumentNullException(nameof(report), Constants.NullParamMessage);
            }
            _report = report;
        }

        public object GetMethodContext(string className)
        {
            if (string.IsNullOrWhiteSpace(className))
            {
                return _report;
            }

            Assembly assembly = Assembly.GetExecutingAssembly();
            // TODO Пока простейший вариант поиска. Тут нужно будет подумать как правильно искать
            // с учётом того, что класс может быть переопределён для компании,
            // TODO сделать кеш объектов
            Type type = assembly.GetTypes().Single(t => t.Name == className);
            if (type == null)
            {
                throw new TypeNotFoundException($"Cannot find type \"{className}\"");
            }

            return Activator.CreateInstance(type);
        }
    }
}