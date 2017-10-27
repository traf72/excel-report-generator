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
        private readonly IDictionary<Type, object> _instanceCache = new Dictionary<Type, object>();

        public MethodCallValueProvider(ITypeProvider typeProvider, object defaultInstance)
        {
            if (typeProvider == null)
            {
                throw new ArgumentNullException(nameof(typeProvider), Constants.NullParamMessage);
            }

            TypeProvider = typeProvider;
            DefaultInstance = defaultInstance;
            if (DefaultInstance != null)
            {
                _instanceCache[DefaultInstance.GetType()] = DefaultInstance;
            }
        }

        protected ITypeProvider TypeProvider { get; }

        protected object DefaultInstance { get; }

        protected string MethodCallTemplate { get; private set; }

        protected ITemplateProcessor TemplateProcessor { get; private set; }

        protected HierarchicalDataItem DataItem { get; private set; }

        protected bool IsStatic { get; private set; }

        public virtual object CallMethod(string methodCallTemplate, ITemplateProcessor templateProcessor, HierarchicalDataItem dataItem, bool isStatic = false)
        {
            if (string.IsNullOrWhiteSpace(methodCallTemplate))
            {
                throw new ArgumentException(Constants.EmptyStringParamMessage, nameof(methodCallTemplate));
            }

            MethodCallTemplate = methodCallTemplate.Trim();
            TemplateProcessor = templateProcessor;
            DataItem = dataItem;
            IsStatic = isStatic;

            MethodCallTemplateParts templateParts = ParseTemplate(MethodCallTemplate);
            Type type = GetType(templateParts.TypeName);
            object instance = GetInstance(type);
            return GetMethod(type, templateParts.MethodName).Invoke(instance, GetParams(templateParts.MethodParams));
        }

        protected virtual MethodCallTemplateParts ParseTemplate(string template)
        {
            string incorrectTemplateMessage = $"Template \"{template}\" is incorrect";
            int firstParenthesisIndex = template.IndexOf('(');
            if (firstParenthesisIndex == -1)
            {
                throw new IncorrectTemplateException(incorrectTemplateMessage);
            }

            int lastParenthesisIndex = template.LastIndexOf(')');
            if (lastParenthesisIndex == -1)
            {
                throw new IncorrectTemplateException(incorrectTemplateMessage);
            }

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

            return new MethodCallTemplateParts(typeName?.Trim(), methodName.Trim(), methodParams.Trim());
        }

        private Type GetType(string typeName)
        {
            return string.IsNullOrWhiteSpace(typeName) ? GetDefaultType() : TypeProvider.GetType(typeName);
        }

        protected virtual Type GetDefaultType()
        {
            if (DefaultInstance == null)
            {
                throw new InvalidOperationException($"Type name is not specified in template \"{MethodCallTemplate}\" but defaultInstance is null");
            }
            return DefaultInstance.GetType();
        }

        protected virtual object GetInstance(Type type)
        {
            if (IsStatic)
            {
                return null;
            }

            object instance;
            if (_instanceCache.TryGetValue(type, out instance))
            {
                return instance;
            }
            instance = Activator.CreateInstance(type);
            _instanceCache[type] = instance;
            return instance;
        }

        protected virtual MethodInfo GetMethod(Type type, string methodName)
        {
            BindingFlags methodTypeBindingFlag = IsStatic ? BindingFlags.Static : BindingFlags.Instance;
            MethodInfo method = type.GetMethod(methodName, BindingFlags.Public | methodTypeBindingFlag | BindingFlags.FlattenHierarchy);
            if (method == null)
            {
                throw new MethodNotFoundException($"Could not find public {(IsStatic ? "static " : string.Empty)}method \"{methodName}\" in type \"{type.Name}\" and all its parents");
            }
            return method;
        }

        private object[] GetParams(string methodParams)
        {
            IList<object> callParams = new List<object>();
            string pattern = GetTemplatePatternWithoutBorders();
            foreach (string p in ParseParams(methodParams))
            {
                if (p.StartsWith("\"") && p.EndsWith("\""))
                {
                    callParams.Add(p.Substring(1, p.Length - 2));
                }
                else if (Regex.IsMatch(p, $@"^{pattern}$"))
                {
                    callParams.Add(TemplateProcessor.GetValue(p, DataItem));
                }
                else
                {
                    callParams.Add(p);
                }
            }
            return callParams.ToArray();
        }

        private string GetTemplatePatternWithoutBorders()
        {
            string pattern = TemplateProcessor.TemplatePattern;
            if (TemplateProcessor.LeftTemplateBorder != null)
            {
                pattern = pattern.Substring(TemplateProcessor.LeftTemplateBorder.Length);
            }
            if (TemplateProcessor.RightTemplateBorder != null)
            {
                pattern = pattern.Substring(0, pattern.Length - TemplateProcessor.RightTemplateBorder.Length);
            }
            return pattern;
        }

        protected virtual string[] ParseParams(string methodParams)
        {
            if (string.IsNullOrWhiteSpace(methodParams))
            {
                return new string[0];
            }

            IList<string> result = new List<string>();
            int currentSymbolIndex = 0;
            int methodNesting = 0;
            StringBuilder param = new StringBuilder();
            while (currentSymbolIndex < methodParams.Length)
            {
                char currentSymbol = methodParams[currentSymbolIndex];
                bool nextSymbolExists = currentSymbolIndex != methodParams.Length - 1;
                char? nextSymbol = nextSymbolExists ? (char?)methodParams[currentSymbolIndex + 1] : null;
                if (currentSymbol == ',')
                {
                    if (nextSymbol == ',' || methodNesting > 0)
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
                    if (currentSymbol == '(')
                    {
                        methodNesting++;
                    }
                    else if (currentSymbol == ')')
                    {
                        methodNesting--;
                    }
                    param.Append(currentSymbol);
                }

                currentSymbolIndex++;
            }

            result.Add(param.ToString());
            return result.Select(p => p.Trim()).ToArray();
        }

        protected class MethodCallTemplateParts
        {
            public MethodCallTemplateParts(string typeName, string methodName, string methodParams)
            {
                TypeName = typeName;
                MethodName = methodName;
                MethodParams = methodParams;
            }

            public string TypeName { get; }

            public string MethodName { get; }

            public string MethodParams { get; }
        }
    }
}