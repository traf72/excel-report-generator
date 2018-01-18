using System;
using ExcelReportGenerator.Rendering.TemplateProcessors;

namespace ExcelReportGenerator.Rendering.Providers
{
    public interface IMethodCallValueProvider
    {
        /// <summary>
        /// Call method by template
        /// </summary>
        /// <param name="templateProcessor">Template processor that will be used for parameters specified as templates</param>
        /// <param name="dataItem">Data item that will be used for parameters specified as data item templates</param>
        /// <returns>Method result</returns>
        object CallMethod(string methodCallTemplate, ITemplateProcessor templateProcessor, HierarchicalDataItem dataItem);

        /// <summary>
        /// Call method by template
        /// </summary>
        /// <param name="concreteType">Concrete type where method will be searched</param>
        /// <param name="templateProcessor">Template processor that will be used for parameters specified as templates</param>
        /// <param name="dataItem">Data item that will be used for parameters specified as data item templates</param>
        /// <returns>Method result</returns>
        object CallMethod(string methodCallTemplate, Type concreteType, ITemplateProcessor templateProcessor, HierarchicalDataItem dataItem);
    }
}