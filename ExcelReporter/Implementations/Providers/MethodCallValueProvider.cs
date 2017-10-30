using ExcelReporter.Exceptions;
using ExcelReporter.Extensions;
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

        private IList<InputParameter> _inputParameters;

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
            FillInputParameters(templateParts.MethodParams);
            return CallMethod(instance, GetMethod(type, templateParts.MethodName));
        }

        private object CallMethod(object instance, MethodInfo method)
        {
            ParameterInfo[] methodParameters = method.GetParameters();
            if (_inputParameters.Count > methodParameters.Length)
            {
                throw new InvalidOperationException($"Mismatch parameters count. Input pararameters count: {_inputParameters.Count}. Method parameters count: {methodParameters.Length}. MethodCallTemplate: {MethodCallTemplate}");
            }

            ParameterInfo[] requiredParams = methodParameters.Where(p => !p.IsOptional).ToArray();
            if (_inputParameters.Count < requiredParams.Length)
            {
                throw new InvalidOperationException($"Mismatch parameters count. Input pararameters count: {_inputParameters.Count}. Method required parameters count: {requiredParams.Length}. MethodCallTemplate: {MethodCallTemplate}");
            }

            object[] callParams = methodParameters.Select(p => p.HasDefaultValue ? p.DefaultValue : null).ToArray();
            for (int i = 0; i < _inputParameters.Count; i++)
            {
                InputParameter param = _inputParameters[i];
                if (param.Type != null || param.Value == null)
                {
                    callParams[i] = param.Value;
                }
                else
                {
                    Type paramType = methodParameters[i].ParameterType;
                    callParams[i] = Convert.ChangeType(param.Value, paramType);
                }
            }
            return method.Invoke(instance, callParams);
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

        private void FillInputParameters(string methodParams)
        {
            _inputParameters = new List<InputParameter>();
            string pattern = GetTemplatePatternWithoutBorders();
            foreach (string p in ParseParams(methodParams))
            {
                if (p.StartsWith("\"") && p.EndsWith("\""))
                {
                    _inputParameters.Add(new InputParameter
                    {
                        Value = p.Substring(1, p.Length - 2),
                        Type = typeof(string),
                    });
                }
                else if (Regex.IsMatch(p, $@"^{pattern}$"))
                {
                    object value = TemplateProcessor.GetValue(p, DataItem);
                    _inputParameters.Add(new InputParameter
                    {
                        Value = value,
                        Type = value?.GetType(),
                    });
                }
                else
                {
                    object value = p;
                    Type type = null;
                    Match match = Regex.Match(p, @"^\[(.+)\](.+)$");
                    if (match.Success)
                    {
                        type = GetTypeByCode(match.Groups[1].Value);
                        value = Convert.ChangeType(match.Groups[2].Value.Trim(), type);
                    }
                    _inputParameters.Add(new InputParameter { Value = value, Type = type });
                }
            }
        }

        private Type GetTypeByCode(string code)
        {
            if (string.IsNullOrWhiteSpace(code))
            {
                return null;
            }

            switch (code.Trim().ToLower())
            {
                case "byte":
                    return typeof(sbyte);

                case "short":
                case "int16":
                    return typeof(short);

                case "int":
                case "int32":
                    return typeof(int);

                case "long":
                case "int64":
                    return typeof(long);

                case "float":
                case "single":
                    return typeof(float);

                case "double":
                    return typeof(double);

                case "decimal":
                    return typeof(decimal);

                case "bool":
                case "boolean":
                    return typeof(bool);

                case "char":
                    return typeof(char);

                case "string":
                    return typeof(string);

                case "datetime":
                case "date":
                    return typeof(DateTime);
            }

            throw new NotSupportedException($"Type \"{code}\" is not supported");
        }

        protected virtual MethodInfo GetMethod(Type type, string methodName)
        {
            string methodNotFoundMessageTemplate = $"Could not find public {(IsStatic ? "static " : string.Empty)}method \"{methodName}\" in type \"{type.Name}\" and all its parents";
            BindingFlags methodTypeBindingFlag = IsStatic ? BindingFlags.Static : BindingFlags.Instance;
            IList<MethodInfo> methods = type.GetMethods(BindingFlags.Public | methodTypeBindingFlag | BindingFlags.FlattenHierarchy).Where(m => m.Name == methodName).ToList();
            if (!methods.Any())
            {
                throw new MethodNotFoundException($"{methodNotFoundMessageTemplate}. MethodCallTemplate: {MethodCallTemplate}");
            }
            if (methods.Any(m => m.GetParameters().Any(p => p.IsParams())))
            {
                throw new NotSupportedException($"Methods which have \"params\" argument are not supported. MethodCallTemplate: {MethodCallTemplate}");
            }
            if (methods.Count == 1)
            {
                return methods.First();
            }

            methods = methods.Where(m =>
            {
                ParameterInfo[] allParams = m.GetParameters();
                IEnumerable<ParameterInfo> optionalParams = allParams.Where(p => p.IsOptional);
                int paramsCountDiff = allParams.Length - _inputParameters.Count;
                return paramsCountDiff >= 0 && paramsCountDiff <= optionalParams.Count();
            }).ToList();

            if (methods.Count == 1)
            {
                return methods.First();
            }
            if (!methods.Any())
            {
                throw new MethodNotFoundException($"{methodNotFoundMessageTemplate} with suitable number of parameters. MethodCallTemplate: {MethodCallTemplate}");
            }
            if (_inputParameters.Any(p => p.Value != null && p.Type == null))
            {
                throw new NotSupportedException($"More than one method found with suitable number of parameters but some of static parameters does not specify a type explicitly. Specify the type explicitly for all static parameters and try again. MethodCallTemplate: {MethodCallTemplate}");
            }

            MethodInfo method = null;
            foreach (MethodInfo m in methods)
            {
                bool isSuitable = true;
                ParameterInfo[] parameters = m.GetParameters();
                for (int i = 0; i < _inputParameters.Count; i++)
                {
                    if (parameters[i].ParameterType != _inputParameters[i].Type)
                    {
                        isSuitable = false;
                        break;
                    }
                }
                if (isSuitable)
                {
                    method = m;
                    break;
                }
            }

            if (method == null)
            {
                throw new NotSupportedException($"More than one method found with suitable number of parameters. In this case the method is chosen by exact match of parameter types. None of methods is suitable. MethodCallTemplate: {MethodCallTemplate}");
            }

            return method;
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

        private class InputParameter
        {
            public object Value { get; set; }

            public Type Type { get; set; }
        }
    }
}