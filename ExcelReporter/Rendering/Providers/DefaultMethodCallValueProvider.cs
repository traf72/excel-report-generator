using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using ExcelReporter.Exceptions;
using ExcelReporter.Extensions;
using ExcelReporter.Helpers;
using ExcelReporter.Rendering.TemplateProcessors;

namespace ExcelReporter.Rendering.Providers
{
    /// <summary>
    /// Provides result of method call
    /// </summary>
    public class DefaultMethodCallValueProvider : IMethodCallValueProvider
    {
        private static readonly Stack<string> TemplateStack = new Stack<string>();

        private readonly IDictionary<Type, object> _instanceCache = new Dictionary<Type, object>();

        /// <param name="typeProvider">Type provider which will be used for type search</param>
        /// <param name="defaultInstance">Instance on which method will be called if template does not specify the type explicitly</param>
        public DefaultMethodCallValueProvider(ITypeProvider typeProvider, object defaultInstance)
        {
            TypeProvider = typeProvider ?? throw new ArgumentNullException(nameof(typeProvider), ArgumentHelper.NullParamMessage);
            DefaultInstance = defaultInstance;
            if (DefaultInstance != null)
            {
                _instanceCache[DefaultInstance.GetType()] = DefaultInstance;
            }
        }

        /// <summary>
        /// Type provider used for type search
        /// </summary>
        protected ITypeProvider TypeProvider { get; }

        /// <summary>
        /// Instance on which method are called if template does not specify the type explicitly
        /// </summary>
        protected object DefaultInstance { get; }

        private string MethodCallTemplate => TemplateStack.Peek();

        /// <summary>
        /// Call method by template
        /// </summary>
        /// <param name="templateProcessor">Template processor that will be used for parameters specified as templates</param>
        /// <param name="dataItem">Data item that will be used for parameters specified as data item templates</param>
        /// <param name="isStatic">Is called method static</param>
        /// <returns>Method result</returns>
        public virtual object CallMethod(string methodCallTemplate, ITemplateProcessor templateProcessor, object dataItem, bool isStatic = false)
        {
            if (string.IsNullOrWhiteSpace(methodCallTemplate))
            {
                throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(methodCallTemplate));
            }

            TemplateStack.Push(methodCallTemplate.Trim());
            try
            {
                MethodCallTemplateParts templateParts = ParseTemplate(MethodCallTemplate);
                Type type = GetType(templateParts.TypeName);
                object instance = GetInstance(type, isStatic);
                IList<InputParameter> inputParams = GetInputParametersValues(templateParts.MethodParams, templateProcessor, dataItem);
                MethodInfo method = GetMethod(type, templateParts.MethodName, inputParams, isStatic);
                return CallMethod(instance, method, inputParams);
            }
            finally
            {
                TemplateStack.Pop();
            }
        }

        /// <param name="instance">Instance on which method will be called</param>
        /// <param name="inputParameters">Parameters with which method will be called</param>
        private object CallMethod(object instance, MethodInfo method, IList<InputParameter> inputParameters)
        {
            ParameterInfo[] methodParameters = method.GetParameters();
            if (inputParameters.Count > methodParameters.Length)
            {
                throw new InvalidOperationException($"Mismatch parameters count. Input pararameters count: {inputParameters.Count}. Method parameters count: {methodParameters.Length}. MethodCallTemplate: {MethodCallTemplate}");
            }

            ParameterInfo[] requiredParams = methodParameters.Where(p => !p.IsOptional).ToArray();
            if (inputParameters.Count < requiredParams.Length)
            {
                throw new InvalidOperationException($"Mismatch parameters count. Input pararameters count: {inputParameters.Count}. Method required parameters count: {requiredParams.Length}. MethodCallTemplate: {MethodCallTemplate}");
            }

            object[] callParams = methodParameters.Select(p => p.HasDefaultValue ? p.DefaultValue : null).ToArray();
            for (int i = 0; i < inputParameters.Count; i++)
            {
                InputParameter param = inputParameters[i];
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

        /// <summary>
        /// Provides the default type where methods are searched if template does not specify the type explicitly
        /// </summary>
        protected virtual Type GetDefaultType()
        {
            if (DefaultInstance == null)
            {
                throw new InvalidOperationException($"Type name is not specified in template \"{MethodCallTemplate}\" but defaultInstance is null");
            }
            return DefaultInstance.GetType();
        }

        /// <summary>
        /// Provides instance on which method will be called
        /// </summary>
        /// <param name="type">Instance type</param>
        protected virtual object GetInstance(Type type, bool isMethodStatic)
        {
            if (isMethodStatic)
            {
                return null;
            }

            if (_instanceCache.TryGetValue(type, out object instance))
            {
                return instance;
            }
            instance = Activator.CreateInstance(type);
            _instanceCache[type] = instance;
            return instance;
        }

        /// <summary>
        /// Parse input parameters string and returns list of input parameters values
        /// </summary>
        /// <param name="templateProcessor">Template processor that will be used for parameters specified as templates</param>
        /// <param name="dataItem">Data item that will be used for parameters specified as data item templates</param>
        protected virtual IList<InputParameter> GetInputParametersValues(string inputParamsAsString, ITemplateProcessor templateProcessor, object dataItem)
        {
            IList<InputParameter> inputParameters = new List<InputParameter>();
            string pattern = templateProcessor.GetTemplateWithoutBorders(templateProcessor.TemplatePattern);
            foreach (string p in ParseInputParams(inputParamsAsString))
            {
                if (p.StartsWith("\"") && p.EndsWith("\""))
                {
                    inputParameters.Add(new InputParameter
                    {
                        Value = p.Substring(1, p.Length - 2),
                        Type = typeof(string),
                    });
                }
                else if (Regex.IsMatch(p, $@"^{pattern}$"))
                {
                    object value = templateProcessor.GetValue(p, dataItem);
                    inputParameters.Add(new InputParameter
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
                    inputParameters.Add(new InputParameter { Value = value, Type = type });
                }
            }
            return inputParameters;
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

        /// <summary>
        /// Find method by name in specified type
        /// </summary>
        /// <param name="type">Type where method will be searched</param>
        /// <param name="inputParameters">List of input method parameters</param>
        /// <param name="isStatic">Is method static</param>
        protected virtual MethodInfo GetMethod(Type type, string methodName, IList<InputParameter> inputParameters, bool isStatic)
        {
            string methodNotFoundMessageTemplate = $"Could not find public {(isStatic ? "static " : string.Empty)}method \"{methodName}\" in type \"{type.Name}\" and all its parents";
            BindingFlags methodTypeBindingFlag = isStatic ? BindingFlags.Static : BindingFlags.Instance;
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
                int paramsCountDiff = allParams.Length - inputParameters.Count;
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
            if (inputParameters.Any(p => p.Value != null && p.Type == null))
            {
                throw new NotSupportedException($"More than one method found with suitable number of parameters but some of static parameters does not specify a type explicitly. Specify the type explicitly for all static parameters and try again. MethodCallTemplate: {MethodCallTemplate}");
            }

            MethodInfo method = null;
            foreach (MethodInfo m in methods)
            {
                bool isSuitable = true;
                ParameterInfo[] parameters = m.GetParameters();
                for (int i = 0; i < inputParameters.Count; i++)
                {
                    if (parameters[i].ParameterType != inputParameters[i].Type)
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

        /// <summary>
        /// Parse input parameters string into array
        /// </summary>
        private string[] ParseInputParams(string inputParamsAsString)
        {
            if (string.IsNullOrWhiteSpace(inputParamsAsString))
            {
                return new string[0];
            }

            IList<string> result = new List<string>();
            int currentSymbolIndex = 0;
            int methodNesting = 0;
            StringBuilder param = new StringBuilder();
            while (currentSymbolIndex < inputParamsAsString.Length)
            {
                char currentSymbol = inputParamsAsString[currentSymbolIndex];
                bool nextSymbolExists = currentSymbolIndex != inputParamsAsString.Length - 1;
                char? nextSymbol = nextSymbolExists ? (char?)inputParamsAsString[currentSymbolIndex + 1] : null;
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

        /// <summary>
        /// Represent parts from which template consist of
        /// </summary>
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

        /// <summary>
        /// Represent input method parameter
        /// </summary>
        protected class InputParameter
        {
            public object Value { get; set; }

            public Type Type { get; set; }
        }
    }
}