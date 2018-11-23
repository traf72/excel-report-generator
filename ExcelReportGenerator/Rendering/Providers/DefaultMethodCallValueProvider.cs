using ExcelReportGenerator.Exceptions;
using ExcelReportGenerator.Extensions;
using ExcelReportGenerator.Helpers;
using ExcelReportGenerator.Rendering.TemplateParts;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelReportGenerator.Rendering.Providers
{
    /// <summary>
    /// Default implementation of <see cref="IMethodCallValueProvider" />
    /// </summary>
    public class DefaultMethodCallValueProvider : IMethodCallValueProvider
    {
        private readonly Stack<string> _templateStack = new Stack<string>();

        /// <param name="typeProvider">Type provider which will be used for type search</param>
        /// <param name="instanceProvider">Instance provider which will be used to get instance of specified type</param>
        public DefaultMethodCallValueProvider(ITypeProvider typeProvider, IInstanceProvider instanceProvider)
        {
            TypeProvider = typeProvider ?? throw new ArgumentNullException(nameof(typeProvider), ArgumentHelper.NullParamMessage);
            InstanceProvider = instanceProvider ?? throw new ArgumentNullException(nameof(instanceProvider), ArgumentHelper.NullParamMessage);
        }

        /// <summary>
        /// Type provider used for type search
        /// </summary>
        protected ITypeProvider TypeProvider { get; }

        /// <summary>
        /// Instance provider used to get instance of specified type
        /// </summary>
        protected IInstanceProvider InstanceProvider { get; }

        private string MethodCallTemplate => _templateStack.Peek();

        /// <inheritdoc />
        /// <seealso cref="CallMethod(string, Type, ITemplateProcessor, HierarchicalDataItem)"/>
        public virtual object CallMethod(string methodCallTemplate, ITemplateProcessor templateProcessor, HierarchicalDataItem dataItem)
        {
            return CallMethod(methodCallTemplate, null, templateProcessor, dataItem);
        }

        /// <inheritdoc />
        /// <exception cref="ArgumentException">Thrown when <paramref name="methodCallTemplate" /> is null, empty string or whitespace</exception>
        /// <exception cref="InvalidTemplateException"></exception>
        /// <exception cref="MethodNotFoundException"></exception>
        /// <exception cref="NotSupportedException"></exception>
        /// <exception cref="InvalidOperationException"></exception>
        public object CallMethod(string methodCallTemplate, Type concreteType, ITemplateProcessor templateProcessor, HierarchicalDataItem dataItem)
        {
            if (string.IsNullOrWhiteSpace(methodCallTemplate))
            {
                throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(methodCallTemplate));
            }

            _templateStack.Push(methodCallTemplate.Trim());
            try
            {
                MethodCallTemplateParts templateParts = ParseTemplate(MethodCallTemplate);
                Type type = concreteType ?? TypeProvider.GetType(templateParts.TypeName);
                IList<InputParameter> inputParams = GetInputParametersValues(templateParts.MethodParams, templateProcessor, dataItem);
                MethodInfo method = GetMethod(type, templateParts.MemberName, inputParams);
                object instance = GetInstance(type, method.IsStatic);
                return CallMethod(instance, method, inputParams);
            }
            finally
            {
                _templateStack.Pop();
            }
        }

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

            object[] callParams = methodParameters.Select(p => p.HasDefaultValue() ? p.DefaultValue : null).ToArray();
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

        /// <summary>
        /// Parse method call <paramref name="template"/> into <see cref="MethodCallTemplateParts"/>
        /// </summary>
        protected virtual MethodCallTemplateParts ParseTemplate(string template)
        {
            string invalidTemplateMessage = string.Format(Constants.InvalidTemplateMessage, template);
            int firstParenthesisIndex = template.IndexOf('(');
            if (firstParenthesisIndex == -1)
            {
                throw new InvalidTemplateException(invalidTemplateMessage);
            }

            int lastParenthesisIndex = template.LastIndexOf(')');
            if (lastParenthesisIndex == -1)
            {
                throw new InvalidTemplateException(invalidTemplateMessage);
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

        // Provides instance of type on which method will be called
        private object GetInstance(Type type, bool isMethodStatic)
        {
            return isMethodStatic ? null : InstanceProvider.GetInstance(type);
        }

        /// <summary>
        /// Parse input parameters string and returns list of input parameters values
        /// </summary>
        /// <param name="templateProcessor">Template processor that will be used for parameters specified as templates</param>
        /// <param name="dataItem">Data item that will be used for parameters specified as data item templates</param>
        protected virtual IList<InputParameter> GetInputParametersValues(string inputParamsAsString, ITemplateProcessor templateProcessor, HierarchicalDataItem dataItem)
        {
            IList<InputParameter> inputParameters = new List<InputParameter>();
            string pattern = templateProcessor.UnwrapTemplate(templateProcessor.GetFullRegexPattern(), true);
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
        protected virtual MethodInfo GetMethod(Type type, string methodName, IList<InputParameter> inputParameters)
        {
            string methodNotFoundMessageTemplate = $"Could not find public method \"{methodName}\" in type \"{type.Name}\" and all its parents";
            IList<MethodInfo> methods = type.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.FlattenHierarchy)
                .Where(m => m.Name == methodName).ToList();

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

        // Parse input parameters string into array
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
        /// Represent input method parameter
        /// </summary>
        protected class InputParameter
        {
            public object Value { get; set; }

            public Type Type { get; set; }
        }
    }
}