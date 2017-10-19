using ExcelReporter.Exceptions;
using ExcelReporter.Interfaces.Providers;
using ExcelReporter.Interfaces.TemplateProcessors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
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

            IList<object> callParams = new List<object>();
            foreach (string p in ParseParams(templateParts.MethodParams))
            {
                if (Regex.IsMatch(p, $@"^{templateProcessor.Pattern}$"))
                {
                    callParams.Add(templateProcessor.GetValue(p, dataItem));
                }
                else
                {
                    callParams.Add(p);
                }
            }

            //object[] callParams = ParseParams(templateParts.MethodParams)
            //    .Select(p => Regex.IsMatch(p, $@"^{templateProcessor.Pattern}$") ? templateProcessor.GetValue(p, dataItem) : p)
            //    .ToArray();

            //object[] callParams = templateParts.MethodParams
            //    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
            //    .Select(p => Regex.IsMatch(p, $@"^{templateProcessor.Pattern}$") ? templateProcessor.GetValue(p.Trim(), dataItem) : p)
            //    .ToArray();

            MethodInfo method = type.GetMethod(templateParts.MethodName);
            return method.Invoke(instance, callParams.ToArray());
        }

        protected virtual MethodCallTemplateParts ParseTemplate(string template)
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

        protected virtual string[] ParseParams(string methodParams)
        {
            if (string.IsNullOrWhiteSpace(methodParams))
            {
                return new string[0];
            }

            IList<string> result = new List<string>();
            int currentSymbolIndex = 0;
            int templateNesting = 0;
            StringBuilder param = new StringBuilder();
            while (currentSymbolIndex < methodParams.Length)
            {
                char currentSymbol = methodParams[currentSymbolIndex];
                bool nextSymbolExists = currentSymbolIndex != methodParams.Length - 1;
                char? nextSymbol = nextSymbolExists ? (char?)methodParams[currentSymbolIndex + 1] : null;
                if (currentSymbol == ',')
                {
                    if (nextSymbol == ',' || templateNesting > 0)
                    {
                        param.Append(currentSymbol);
                        if (nextSymbol == ',')
                        {
                            currentSymbolIndex++;
                        }
                    }
                    else
                    {
                        result.Add(param.ToString());
                        param.Clear();
                    }
                }
                else
                {
                    if (currentSymbol == '{')
                    {
                        templateNesting++;
                    }
                    else if (currentSymbol == '}')
                    {
                        templateNesting--;
                    }
                    param.Append(currentSymbol);
                }

                currentSymbolIndex++;
            }

            result.Add(param.ToString());
            return result.Select(p => p.Trim()).ToArray();
        }

        protected virtual object CreateInstance(Type type)
        {
            return Activator.CreateInstance(type);
        }
    }

    public class MethodCallTemplateParts
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