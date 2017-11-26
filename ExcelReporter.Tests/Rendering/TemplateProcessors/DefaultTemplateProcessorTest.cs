using System;
using ExcelReporter.Exceptions;
using ExcelReporter.Rendering;
using ExcelReporter.Rendering.Providers;
using ExcelReporter.Rendering.Providers.DataItemValueProviders;
using ExcelReporter.Rendering.Providers.ParameterProviders;
using ExcelReporter.Rendering.TemplateProcessors;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;

namespace ExcelReporter.Tests.Rendering.TemplateProcessors
{
    [TestClass]
    public class DefaultTemplateProcessorTest
    {
        [TestMethod]
        public void TestGetValue()
        {
            var parameterProvider = Substitute.For<IParameterProvider>();
            var methodCallValueProvider = Substitute.For<IMethodCallValueProvider>();
            var dataItemValueProvider = Substitute.For<IGenericDataItemValueProvider<HierarchicalDataItem>>();
            var dataItem = new HierarchicalDataItem();

            ExceptionAssert.Throws<ArgumentNullException>(() => new DefaultTemplateProcessor(null, methodCallValueProvider, dataItemValueProvider));
            new DefaultTemplateProcessor(parameterProvider, null, dataItemValueProvider);
            new DefaultTemplateProcessor(parameterProvider, methodCallValueProvider);

            var templateProcessor = new DefaultTemplateProcessor(parameterProvider, methodCallValueProvider, dataItemValueProvider);

            ExceptionAssert.Throws<ArgumentNullException>(() => templateProcessor.GetValue(null, dataItem));
            ExceptionAssert.Throws<IncorrectTemplateException>(() => templateProcessor.GetValue("{p-Name}"), "Incorrect template \"{p-Name}\". Cannot find separator \":\" between member type and member template");
            ExceptionAssert.Throws<IncorrectTemplateException>(() => templateProcessor.GetValue("{bad:Name}"), "Incorrect template \"{bad:Name}\". Unknown member type \"bad\"");

            templateProcessor.GetValue("{p:Name}");

            parameterProvider.Received(1).GetParameterValue("Name");
            methodCallValueProvider.DidNotReceiveWithAnyArgs().CallMethod(null, null, null);
            dataItemValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>(), null);

            parameterProvider.ClearReceivedCalls();
            templateProcessor.GetValue(" { p : Name } ");

            parameterProvider.Received(1).GetParameterValue("Name");
            methodCallValueProvider.DidNotReceiveWithAnyArgs().CallMethod(null, null, null);
            dataItemValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>(), null);

            parameterProvider.ClearReceivedCalls();
            templateProcessor.GetValue(" p : Name ");

            parameterProvider.Received(1).GetParameterValue("Name");
            methodCallValueProvider.DidNotReceiveWithAnyArgs().CallMethod(null, null, null);
            dataItemValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>(), null);

            parameterProvider.ClearReceivedCalls();
            templateProcessor.GetValue("{ m:Method() }", dataItem);

            parameterProvider.DidNotReceiveWithAnyArgs().GetParameterValue(Arg.Any<string>());
            methodCallValueProvider.Received(1).CallMethod("Method()", templateProcessor, dataItem);
            dataItemValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>(), null);

            methodCallValueProvider.ClearReceivedCalls();
            templateProcessor.GetValue("{ms:Method()}", dataItem);
            parameterProvider.DidNotReceiveWithAnyArgs().GetParameterValue(Arg.Any<string>());
            methodCallValueProvider.Received(1).CallMethod("Method()", templateProcessor, dataItem, true);
            dataItemValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>(), null);

            methodCallValueProvider.ClearReceivedCalls();
            ExceptionAssert.Throws<InvalidOperationException>(() => new DefaultTemplateProcessor(parameterProvider).GetValue("{ms:Method()}", null), "Template \"{ms:Method()}\" contains method call but methodCallValueProvider is null");

            templateProcessor.GetValue("{di:Field}", dataItem);
            parameterProvider.DidNotReceiveWithAnyArgs().GetParameterValue(Arg.Any<string>());
            methodCallValueProvider.DidNotReceiveWithAnyArgs().CallMethod(null, null, null);
            dataItemValueProvider.Received(1).GetValue("Field", dataItem);

            dataItemValueProvider.ClearReceivedCalls();
            ExceptionAssert.Throws<InvalidOperationException>(() => templateProcessor.GetValue("{di:Field}"), "Template \"{di:Field}\" contains data reference but dataItem is null");
            ExceptionAssert.Throws<InvalidOperationException>(() => new DefaultTemplateProcessor(parameterProvider).GetValue("{di:Field}", dataItem), "Template \"{di:Field}\" contains data reference but dataItemValueProvider is null");
        }
    }
}