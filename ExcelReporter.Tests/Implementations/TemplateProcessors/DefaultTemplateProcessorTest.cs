using ExcelReporter.Exceptions;
using ExcelReporter.Implementations.TemplateProcessors;
using ExcelReporter.Interfaces.Providers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;
using System;

namespace ExcelReporter.Tests.Implementations.TemplateProcessors
{
    [TestClass]
    public class DefaultTemplateProcessorTest
    {
        [TestMethod]
        public void TestGetValue()
        {
            var parameterProvider = Substitute.For<IParameterProvider>();
            var methodCallValueProvider = Substitute.For<IMethodCallValueProvider>();
            var dataItemValueProvider = Substitute.For<IDataItemValueProvider>();
            var dataItem = new HierarchicalDataItem();

            MyAssert.Throws<ArgumentNullException>(() => new DefaultTemplateProcessor(null, methodCallValueProvider, dataItemValueProvider));
            new DefaultTemplateProcessor(parameterProvider, null, dataItemValueProvider);
            new DefaultTemplateProcessor(parameterProvider, methodCallValueProvider);

            var templateProcessor = new DefaultTemplateProcessor(parameterProvider, methodCallValueProvider, dataItemValueProvider);

            MyAssert.Throws<ArgumentNullException>(() => templateProcessor.GetValue(null, dataItem));
            MyAssert.Throws<IncorrectTemplateException>(() => templateProcessor.GetValue("{p-Name}", null), "Incorrect template \"{p-Name}\". Cannot find separator \":\" between member type and member template");
            MyAssert.Throws<IncorrectTemplateException>(() => templateProcessor.GetValue("{bad:Name}", null), "Incorrect template \"{bad:Name}\". Unknown member type \"bad\"");

            templateProcessor.GetValue("{p:Name}", null);

            parameterProvider.Received(1).GetParameterValue("Name");
            methodCallValueProvider.DidNotReceiveWithAnyArgs().CallMethod(Arg.Any<string>(), null, null);
            dataItemValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>(), null);

            parameterProvider.ClearReceivedCalls();
            templateProcessor.GetValue(" { p : Name } ", null);

            parameterProvider.Received(1).GetParameterValue("Name");
            methodCallValueProvider.DidNotReceiveWithAnyArgs().CallMethod(Arg.Any<string>(), null, null);
            dataItemValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>(), null);

            parameterProvider.ClearReceivedCalls();
            templateProcessor.GetValue(" p : Name ", null);

            parameterProvider.Received(1).GetParameterValue("Name");
            methodCallValueProvider.DidNotReceiveWithAnyArgs().CallMethod(Arg.Any<string>(), null, null);
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
            MyAssert.Throws<InvalidOperationException>(() => new DefaultTemplateProcessor(parameterProvider).GetValue("{ms:Method()}", null), "Template \"{ms:Method()}\" contains method call but methodCallValueProvider is null");

            templateProcessor.GetValue("{di:Field}", dataItem);
            parameterProvider.DidNotReceiveWithAnyArgs().GetParameterValue(Arg.Any<string>());
            methodCallValueProvider.DidNotReceiveWithAnyArgs().CallMethod(Arg.Any<string>(), null, null);
            dataItemValueProvider.Received(1).GetValue("Field", dataItem);

            dataItemValueProvider.ClearReceivedCalls();
            MyAssert.Throws<InvalidOperationException>(() => templateProcessor.GetValue("{di:Field}", null), "Template \"{di:Field}\" contains data reference but dataItem is null");
            MyAssert.Throws<InvalidOperationException>(() => new DefaultTemplateProcessor(parameterProvider).GetValue("{di:Field}", dataItem), "Template \"{di:Field}\" contains data reference but dataItemValueProvider is null");
        }
    }
}