using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Helpers;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelReportGenerator.Extensions
{
    // TODO Пока отключил все проверки на null и пустые строки, так как это будет влиять на производительность
    // TODO Вместо этого лучше проверять это каким-то образом вначале выполнения отчёта один раз
    internal static class TemplateProcessorExtensions
    {
        private static readonly string[] AllAggregationFuncs = Enum.GetNames(typeof(AggregateFunction)).Where(n => n != AggregateFunction.NoAggregation.ToString()).ToArray();

        /// <summary>
        /// Remove template borders
        /// </summary>
        public static string UnwrapTemplate(this ITemplateProcessor processor, string template, bool isRegex = false)
        {
            //CheckForNullOrWhiteSpace(processor.LeftTemplateBorder, nameof(processor.LeftTemplateBorder));
            //CheckForNullOrWhiteSpace(processor.RightTemplateBorder, nameof(processor.RightTemplateBorder));
            if (template == null)
            {
                throw new ArgumentNullException(nameof(template), ArgumentHelper.NullParamMessage);
            }

            string leftBorder = isRegex ? Regex.Escape(processor.LeftTemplateBorder) : processor.LeftTemplateBorder;
            if (template.StartsWith(leftBorder))
            {
                template = template.Substring(leftBorder.Length);
            }

            string rightBorder = isRegex ? Regex.Escape(processor.RightTemplateBorder) : processor.RightTemplateBorder;
            if (template.EndsWith(rightBorder))
            {
                template = template.Substring(0, template.Length - rightBorder.Length);
            }
            return template;
        }

        /// <summary>
        /// Wrap template with borders
        /// </summary>
        public static string WrapTemplate(this ITemplateProcessor processor, string template, bool isRegex = false)
        {
            //CheckForNullOrWhiteSpace(processor.LeftTemplateBorder, nameof(processor.LeftTemplateBorder));
            //CheckForNullOrWhiteSpace(processor.RightTemplateBorder, nameof(processor.RightTemplateBorder));
            string leftBorder = isRegex ? Regex.Escape(processor.LeftTemplateBorder) : processor.LeftTemplateBorder;
            string rightBorder = isRegex ? Regex.Escape(processor.RightTemplateBorder) : processor.RightTemplateBorder;
            return $"{leftBorder}{template}{rightBorder}";
        }

        public static string BuildPropertyTemplate(this ITemplateProcessor processor, string propertyTemplate)
        {
            //CheckForNullOrWhiteSpace(processor.PropertyMemberLabel, nameof(processor.PropertyMemberLabel));
            return BuildTemplate(processor, processor.PropertyMemberLabel, propertyTemplate);
        }

        public static string BuildDataItemTemplate(this ITemplateProcessor processor, string dataItemTemplate)
        {
            //CheckForNullOrWhiteSpace(processor.DataItemMemberLabel, nameof(processor.DataItemMemberLabel));
            return BuildTemplate(processor, processor.DataItemMemberLabel, dataItemTemplate);
        }

        public static string BuildMethodCallTemplate(this ITemplateProcessor processor, string methodCallTemplate)
        {
            //CheckForNullOrWhiteSpace(processor.MethodCallMemberLabel, nameof(processor.MethodCallMemberLabel));
            return BuildTemplate(processor, processor.MethodCallMemberLabel, methodCallTemplate);
        }

        public static string BuildVariableTemplate(this ITemplateProcessor processor, string variableTemplate)
        {
            //CheckForNullOrWhiteSpace(processor.VariableMemberLabel, nameof(processor.VariableMemberLabel));
            return BuildTemplate(processor, processor.VariableMemberLabel, variableTemplate);
        }

        public static string BuildSystemFunctionTemplate(this ITemplateProcessor processor, string systemFunctionTemplate)
        {
            //CheckForNullOrWhiteSpace(processor.SystemFunctionMemberLabel, nameof(processor.SystemFunctionMemberLabel));
            return BuildTemplate(processor, processor.SystemFunctionMemberLabel, systemFunctionTemplate);
        }

        private static string BuildTemplate(ITemplateProcessor processor, string memberLabel, string memberTemplate)
        {
            //CheckForNullOrWhiteSpace(processor.LeftTemplateBorder, nameof(processor.LeftTemplateBorder));
            //CheckForNullOrWhiteSpace(processor.RightTemplateBorder, nameof(processor.RightTemplateBorder));
            //CheckForNullOrEmpty(processor.MemberLabelSeparator, nameof(processor.MemberLabelSeparator));
            return $@"{processor.LeftTemplateBorder}{memberLabel}{processor.MemberLabelSeparator}{memberTemplate}{processor.RightTemplateBorder}";
        }

        public static string TrimPropertyLabel(this ITemplateProcessor processor, string propertyTemplate)
        {
            //CheckForNullOrWhiteSpace(processor.PropertyMemberLabel, nameof(processor.PropertyMemberLabel));
            return TrimMemberLabel(processor, processor.PropertyMemberLabel, propertyTemplate);
        }

        public static string TrimDataItemLabel(this ITemplateProcessor processor, string dataItemTemplate)
        {
            //CheckForNullOrWhiteSpace(processor.DataItemMemberLabel, nameof(processor.DataItemMemberLabel));
            return TrimMemberLabel(processor, processor.DataItemMemberLabel, dataItemTemplate);
        }

        public static string TrimMethodCallLabel(this ITemplateProcessor processor, string methodCallTemplate)
        {
            //CheckForNullOrWhiteSpace(processor.MethodCallMemberLabel, nameof(processor.MethodCallMemberLabel));
            return TrimMemberLabel(processor, processor.MethodCallMemberLabel, methodCallTemplate);
        }

        public static string TrimVariableLabel(this ITemplateProcessor processor, string variableTemplate)
        {
            //CheckForNullOrWhiteSpace(processor.VariableMemberLabel, nameof(processor.VariableMemberLabel));
            return TrimMemberLabel(processor, processor.VariableMemberLabel, variableTemplate);
        }

        public static string TrimSystemFunctionLabel(this ITemplateProcessor processor, string systemFunctionTemplate)
        {
            //CheckForNullOrWhiteSpace(processor.SystemFunctionMemberLabel, nameof(processor.SystemFunctionMemberLabel));
            return TrimMemberLabel(processor, processor.SystemFunctionMemberLabel, systemFunctionTemplate);
        }

        private static string TrimMemberLabel(ITemplateProcessor processor, string memberLabel, string memberTemplate)
        {
            //CheckForNullOrEmpty(processor.MemberLabelSeparator, nameof(processor.MemberLabelSeparator));
            if (memberTemplate == null)
            {
                throw new ArgumentNullException(nameof(memberTemplate), ArgumentHelper.NullParamMessage);
            }

            string memberLabelWithSeparator = $"{memberLabel}{processor.MemberLabelSeparator}";
            int index = memberTemplate.IndexOf(memberLabelWithSeparator, StringComparison.CurrentCultureIgnoreCase);
            if (index == -1)
            {
                return memberTemplate;
            }

            string firstPart = memberTemplate.Substring(0, index);
            string lastPart = memberTemplate.Substring(index + memberLabelWithSeparator.Length);
            return $"{firstPart}{lastPart}";
        }

        /// <summary>
        /// Regex pattern that includes all member types
        /// </summary>
        public static string GetFullRegexPattern(this ITemplateProcessor processor)
        {
            //CheckForNullOrWhiteSpace(processor.PropertyMemberLabel, nameof(processor.PropertyMemberLabel));
            //CheckForNullOrWhiteSpace(processor.DataItemMemberLabel, nameof(processor.DataItemMemberLabel));
            //CheckForNullOrWhiteSpace(processor.MethodCallMemberLabel, nameof(processor.MethodCallMemberLabel));
            //CheckForNullOrWhiteSpace(processor.VariableMemberLabel, nameof(processor.VariableMemberLabel));
            //CheckForNullOrWhiteSpace(processor.SystemFunctionMemberLabel, nameof(processor.SystemFunctionMemberLabel));
            return GetRegexPattern(processor, $"({Regex.Escape(processor.PropertyMemberLabel)}" +
                                              $"|{Regex.Escape(processor.DataItemMemberLabel)}" +
                                              $"|{Regex.Escape(processor.MethodCallMemberLabel)}" +
                                              $"|{Regex.Escape(processor.VariableMemberLabel)}" +
                                              $"|{Regex.Escape(processor.SystemFunctionMemberLabel)})");
        }

        public static string GetPropertyRegexPattern(this ITemplateProcessor processor)
        {
            //CheckForNullOrWhiteSpace(processor.PropertyMemberLabel, nameof(processor.PropertyMemberLabel));
            return GetRegexPattern(processor, $"{Regex.Escape(processor.PropertyMemberLabel)}");
        }

        public static string GetDataItemRegexPattern(this ITemplateProcessor processor)
        {
            //CheckForNullOrWhiteSpace(processor.DataItemMemberLabel, nameof(processor.DataItemMemberLabel));
            return GetRegexPattern(processor, $"{Regex.Escape(processor.DataItemMemberLabel)}");
        }

        public static string GetMethodCallRegexPattern(this ITemplateProcessor processor)
        {
            //CheckForNullOrWhiteSpace(processor.MethodCallMemberLabel, nameof(processor.MethodCallMemberLabel));
            return GetRegexPattern(processor, $"{Regex.Escape(processor.MethodCallMemberLabel)}");
        }

        public static string GetVariableRegexPattern(this ITemplateProcessor processor)
        {
            //CheckForNullOrWhiteSpace(processor.VariableMemberLabel, nameof(processor.VariableMemberLabel));
            return GetRegexPattern(processor, $"{Regex.Escape(processor.VariableMemberLabel)}");
        }

        public static string GetSystemFunctionRegexPattern(this ITemplateProcessor processor)
        {
            //CheckForNullOrWhiteSpace(processor.SystemFunctionMemberLabel, nameof(processor.SystemFunctionMemberLabel));
            return GetRegexPattern(processor, $"{Regex.Escape(processor.SystemFunctionMemberLabel)}");
        }

        private static string GetRegexPattern(ITemplateProcessor processor, string escapedMemberLabel)
        {
            //CheckForNullOrWhiteSpace(processor.LeftTemplateBorder, nameof(processor.LeftTemplateBorder));
            //CheckForNullOrWhiteSpace(processor.RightTemplateBorder, nameof(processor.RightTemplateBorder));
            //CheckForNullOrEmpty(processor.MemberLabelSeparator, nameof(processor.MemberLabelSeparator));
            return $@"{Regex.Escape(processor.LeftTemplateBorder)}\s*{escapedMemberLabel}{Regex.Escape(processor.MemberLabelSeparator)}.+?\s*{Regex.Escape(processor.RightTemplateBorder)}";
        }

        public static string GetFullAggregationRegexPattern(this ITemplateProcessor processor)
        {
            return $@"{Regex.Escape(processor.LeftTemplateBorder)}\s*({string.Join("|", AllAggregationFuncs)})\((.+?)\)\s*{Regex.Escape(processor.RightTemplateBorder)}";
        }

        public static string BuildAggregationFuncTemplate(this ITemplateProcessor processor, AggregateFunction aggFunc, string columnName)
        {
            return $"{processor.LeftTemplateBorder}{aggFunc.ToString()}({processor.DataItemMemberLabel}{processor.MemberLabelSeparator}{columnName}){processor.RightTemplateBorder}";
        }

        //private static void CheckForNullOrWhiteSpace(string value, string propName)
        //{
        //    if (string.IsNullOrWhiteSpace(value))
        //    {
        //        throw new Exception($"{propName} cannot be null, empty or white space");
        //    }
        //}

        //private static void CheckForNullOrEmpty(string value, string propName)
        //{
        //    if (string.IsNullOrEmpty(value))
        //    {
        //        throw new Exception($"{propName} cannot be null or empty");
        //    }
        //}
    }
}