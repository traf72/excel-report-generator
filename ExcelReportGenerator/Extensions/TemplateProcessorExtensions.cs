using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Helpers;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelReportGenerator.Extensions
{
    internal static class TemplateProcessorExtensions
    {
        private static readonly string[] AllAggregationFuncs = Enum.GetNames(typeof(AggregateFunction)).Where(n => n != AggregateFunction.NoAggregation.ToString()).ToArray();

        // Remove template borders
        public static string UnwrapTemplate(this ITemplateProcessor processor, string template, bool isRegex = false)
        {
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

        // Wrap template with borders
        public static string WrapTemplate(this ITemplateProcessor processor, string template, bool isRegex = false)
        {
            string leftBorder = isRegex ? Regex.Escape(processor.LeftTemplateBorder) : processor.LeftTemplateBorder;
            string rightBorder = isRegex ? Regex.Escape(processor.RightTemplateBorder) : processor.RightTemplateBorder;
            return $"{leftBorder}{template}{rightBorder}";
        }

        public static string BuildPropertyTemplate(this ITemplateProcessor processor, string propertyTemplate)
        {
            return BuildTemplate(processor, processor.PropertyMemberLabel, propertyTemplate);
        }

        public static string BuildDataItemTemplate(this ITemplateProcessor processor, string dataItemTemplate)
        {
            return BuildTemplate(processor, processor.DataItemMemberLabel, dataItemTemplate);
        }

        public static string BuildMethodCallTemplate(this ITemplateProcessor processor, string methodCallTemplate)
        {
            return BuildTemplate(processor, processor.MethodCallMemberLabel, methodCallTemplate);
        }

        public static string BuildVariableTemplate(this ITemplateProcessor processor, string variableTemplate)
        {
            return BuildTemplate(processor, processor.SystemVariableMemberLabel, variableTemplate);
        }

        public static string BuildSystemFunctionTemplate(this ITemplateProcessor processor, string systemFunctionTemplate)
        {
            return BuildTemplate(processor, processor.SystemFunctionMemberLabel, systemFunctionTemplate);
        }

        private static string BuildTemplate(ITemplateProcessor processor, string memberLabel, string memberTemplate)
        {
            return $@"{processor.LeftTemplateBorder}{memberLabel}{processor.MemberLabelSeparator}{memberTemplate}{processor.RightTemplateBorder}";
        }

        public static string TrimPropertyLabel(this ITemplateProcessor processor, string propertyTemplate)
        {
            return TrimMemberLabel(processor, processor.PropertyMemberLabel, propertyTemplate);
        }

        public static string TrimDataItemLabel(this ITemplateProcessor processor, string dataItemTemplate)
        {
            return TrimMemberLabel(processor, processor.DataItemMemberLabel, dataItemTemplate);
        }

        public static string TrimMethodCallLabel(this ITemplateProcessor processor, string methodCallTemplate)
        {
            return TrimMemberLabel(processor, processor.MethodCallMemberLabel, methodCallTemplate);
        }

        public static string TrimVariableLabel(this ITemplateProcessor processor, string variableTemplate)
        {
            return TrimMemberLabel(processor, processor.SystemVariableMemberLabel, variableTemplate);
        }

        public static string TrimSystemFunctionLabel(this ITemplateProcessor processor, string systemFunctionTemplate)
        {
            return TrimMemberLabel(processor, processor.SystemFunctionMemberLabel, systemFunctionTemplate);
        }

        private static string TrimMemberLabel(ITemplateProcessor processor, string memberLabel, string memberTemplate)
        {
            if (memberTemplate == null)
            {
                throw new ArgumentNullException(nameof(memberTemplate), ArgumentHelper.NullParamMessage);
            }

            string memberLabelWithSeparatorPattern = $@"(.*?){Regex.Escape(memberLabel)}\s*{Regex.Escape(processor.MemberLabelSeparator)}(.*)";
            Match match = Regex.Match(memberTemplate, memberLabelWithSeparatorPattern, RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                return memberTemplate;
            }
            return match.Groups[1].Value + match.Groups[2].Value;
        }

        public static bool IsHorizontalPageBreak(this ITemplateProcessor processor, string input)
        {
            return IsPageBreak(processor, processor.HorizontalPageBreakLabel, input);
        }

        public static bool IsVerticalPageBreak(this ITemplateProcessor processor, string input)
        {
            return IsPageBreak(processor, processor.VerticalPageBreakLabel, input);
        }

        private static bool IsPageBreak(ITemplateProcessor processor, string pageBreakLabel, string input)
        {
            return input != null && Regex.IsMatch(input, $@"^\s*{Regex.Escape(processor.LeftTemplateBorder)}\s*{Regex.Escape(pageBreakLabel)}\s*{Regex.Escape(processor.RightTemplateBorder)}\s*$", RegexOptions.IgnoreCase);
        }

        // Regex pattern that includes all member types
        public static string GetFullRegexPattern(this ITemplateProcessor processor)
        {
            return GetRegexPattern(processor, $"({Regex.Escape(processor.PropertyMemberLabel)}" +
                                              $"|{Regex.Escape(processor.DataItemMemberLabel)}" +
                                              $"|{Regex.Escape(processor.MethodCallMemberLabel)}" +
                                              $"|{Regex.Escape(processor.SystemVariableMemberLabel)}" +
                                              $"|{Regex.Escape(processor.SystemFunctionMemberLabel)})");
        }

        public static string GetPropertyRegexPattern(this ITemplateProcessor processor)
        {
            return GetRegexPattern(processor, $"{Regex.Escape(processor.PropertyMemberLabel)}");
        }

        public static string GetDataItemRegexPattern(this ITemplateProcessor processor)
        {
            return GetRegexPattern(processor, $"{Regex.Escape(processor.DataItemMemberLabel)}");
        }

        public static string GetMethodCallRegexPattern(this ITemplateProcessor processor)
        {
            return GetRegexPattern(processor, $"{Regex.Escape(processor.MethodCallMemberLabel)}");
        }

        public static string GetVariableRegexPattern(this ITemplateProcessor processor)
        {
            return GetRegexPattern(processor, $"{Regex.Escape(processor.SystemVariableMemberLabel)}");
        }

        public static string GetSystemFunctionRegexPattern(this ITemplateProcessor processor)
        {
            return GetRegexPattern(processor, $"{Regex.Escape(processor.SystemFunctionMemberLabel)}");
        }

        private static string GetRegexPattern(ITemplateProcessor processor, string escapedMemberLabel)
        {
            return $@"{Regex.Escape(processor.LeftTemplateBorder)}\s*{escapedMemberLabel}{Regex.Escape(processor.MemberLabelSeparator)}.+?\s*{Regex.Escape(processor.RightTemplateBorder)}";
        }

        public static string GetTemplatesWithAggregationRegexPattern(this ITemplateProcessor processor)
        {
            return $@"{Regex.Escape(processor.LeftTemplateBorder)}[^{Regex.Escape(processor.RightTemplateBorder)}]*({string.Join("|", AllAggregationFuncs)})\(\s*{Regex.Escape(processor.DataItemMemberLabel)}\s*{Regex.Escape(processor.MemberLabelSeparator)}.+?\)[^{Regex.Escape(processor.RightTemplateBorder)}]*{Regex.Escape(processor.RightTemplateBorder)}";
        }

        public static string GetAggregationFuncRegexPattern(this ITemplateProcessor processor)
        {
            return $@"({string.Join("|", AllAggregationFuncs)})\((\s*{Regex.Escape(processor.DataItemMemberLabel)}\s*{Regex.Escape(processor.MemberLabelSeparator)}.+?)\)";
        }

        public static string BuildAggregationFuncTemplate(this ITemplateProcessor processor, AggregateFunction aggFunc, string columnName)
        {
            return $"{processor.LeftTemplateBorder}{aggFunc.ToString()}({processor.DataItemMemberLabel}{processor.MemberLabelSeparator}{columnName}){processor.RightTemplateBorder}";
        }
    }
}