using ExcelReporter.Exceptions;
using ExcelReporter.Interfaces.Providers;
using ExcelReporter.Interfaces.TemplateProcessors;
using System;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace ExcelReporter.Implementations.Providers
{
    public class MethodCallValueProvider : IMethodCallValueProvider
    {
        private readonly ITypeProvider _typeProvider;

        public MethodCallValueProvider(ITypeProvider typeProvider)
        {
            if (typeProvider == null)
            {
                throw new ArgumentNullException(nameof(typeProvider), Constants.NullParamMessage);
            }
            _typeProvider = typeProvider;
        }

        public object CallMethod(string methodCallTemplate, ITemplateProcessor templateProcessor, HierarchicalDataItem dataItem, bool isStatic = false)
        {
            if (string.IsNullOrWhiteSpace(methodCallTemplate))
            {
                throw new ArgumentException(Constants.EmptyStringParamMessage, nameof(methodCallTemplate));
            }

            MethodCallTemplateParts templateParts = ParseTemplate(methodCallTemplate);

            Type type = _typeProvider.GetType(templateParts.TypeName);
            object instance = null;
            if (!isStatic)
            {
                instance = CreateInstance(type);
            }

            object[] callParams = templateParts.MethodParams
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(p => Regex.IsMatch(p, $@"^{templateProcessor.Pattern}$") ? templateProcessor.GetValue(p.Trim(), dataItem) : p)
                .ToArray();

            MethodInfo method = type.GetMethod(templateParts.MethodName);
            return method.Invoke(instance, callParams);
        }

        private MethodCallTemplateParts ParseTemplate(string template)
        {
            int parenthesisCount = Regex.Matches(template, @"\(|\)").Count;
            if (parenthesisCount == 0 || parenthesisCount % 2 != 0)
            {
                throw new IncorrectTemplateException($"Template \"{template}\" is incorrect");
            }

            int firstParenthesisIndex = template.IndexOf('(');
            int lastParenthesisIndex = template.LastIndexOf(')');
            string methodParams = template.Substring(firstParenthesisIndex + 1, lastParenthesisIndex - firstParenthesisIndex - 1);
            string typeWithMethodName = template.Substring(0, firstParenthesisIndex);
            string methodName;
            string typeName = null;
            int lastColonIndex = typeWithMethodName.LastIndexOf(':');
            if (lastColonIndex == -1)
            {
                methodName = typeWithMethodName;
            }
            else
            {
                methodName = typeWithMethodName.Substring(lastColonIndex + 1);
                typeName = typeWithMethodName.Substring(0, lastColonIndex);
            }

            return new MethodCallTemplateParts(typeName, methodName, methodParams);
        }

        protected virtual object CreateInstance(Type type)
        {
            return Activator.CreateInstance(type);
        }
    }

    internal class MethodCallTemplateParts
    {
        public MethodCallTemplateParts(string typeName, string methodName, string methodParams)
        {
            TypeName = typeName;
            MethodName = methodName;
            MethodParams = methodParams;
        }

        public string TypeName { get; set; }

        public string MethodName { get; set; }

        public string MethodParams { get; set; }
    }
}