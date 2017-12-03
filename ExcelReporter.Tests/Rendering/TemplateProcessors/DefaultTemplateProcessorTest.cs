using System;
using ExcelReporter.Exceptions;
using ExcelReporter.Rendering;
using ExcelReporter.Rendering.Providers;
using ExcelReporter.Rendering.Providers.DataItemValueProviders;
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
            var propertyValueProvider = Substitute.For<IPropertyValueProvider>();
            var methodCallValueProvider = Substitute.For<IMethodCallValueProvider>();
            var dataItemValueProvider = Substitute.For<IGenericDataItemValueProvider<HierarchicalDataItem>>();
            var dataItem = new HierarchicalDataItem();

            ExceptionAssert.Throws<ArgumentNullException>(() => new DefaultTemplateProcessor(null, methodCallValueProvider, dataItemValueProvider));
            new DefaultTemplateProcessor(propertyValueProvider, null, dataItemValueProvider);
            new DefaultTemplateProcessor(propertyValueProvider, methodCallValueProvider);

            var templateProcessor = new DefaultTemplateProcessor(propertyValueProvider, methodCallValueProvider, dataItemValueProvider);

            ExceptionAssert.Throws<ArgumentNullException>(() => templateProcessor.GetValue(null, dataItem));
            ExceptionAssert.Throws<IncorrectTemplateException>(() => templateProcessor.GetValue("{p-Name}"), "Incorrect template \"{p-Name}\". Cannot find separator \":\" between member type and member template");
            ExceptionAssert.Throws<IncorrectTemplateException>(() => templateProcessor.GetValue("{bad:Name}"), "Incorrect template \"{bad:Name}\". Unknown member type \"bad\"");

            templateProcessor.GetValue("{p:Name}");

            propertyValueProvider.Received(1).GetValue("Name");
            methodCallValueProvider.DidNotReceiveWithAnyArgs().CallMethod(null, null, null);
            dataItemValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>(), null);

            propertyValueProvider.ClearReceivedCalls();
            templateProcessor.GetValue(" { p : Name } ");

            propertyValueProvider.Received(1).GetValue("Name");
            methodCallValueProvider.DidNotReceiveWithAnyArgs().CallMethod(null, null, null);
            dataItemValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>(), null);

            propertyValueProvider.ClearReceivedCalls();
            templateProcessor.GetValue(" p : Name ");

            propertyValueProvider.Received(1).GetValue("Name");
            methodCallValueProvider.DidNotReceiveWithAnyArgs().CallMethod(null, null, null);
            dataItemValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>(), null);

            propertyValueProvider.ClearReceivedCalls();
            templateProcessor.GetValue("{ m:Method() }", dataItem);

            propertyValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());
            methodCallValueProvider.Received(1).CallMethod("Method()", templateProcessor, dataItem);
            dataItemValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>(), null);

            methodCallValueProvider.ClearReceivedCalls();
            templateProcessor.GetValue("{m:Method()}", dataItem);
            propertyValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());
            methodCallValueProvider.Received(1).CallMethod("Method()", templateProcessor, dataItem);
            dataItemValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>(), null);

            methodCallValueProvider.ClearReceivedCalls();
            ExceptionAssert.Throws<InvalidOperationException>(() => new DefaultTemplateProcessor(propertyValueProvider).GetValue("{m:Method()}"), "Template \"{m:Method()}\" contains method call but methodCallValueProvider is null");

            templateProcessor.GetValue("{di:Field}", dataItem);
            propertyValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());
            methodCallValueProvider.DidNotReceiveWithAnyArgs().CallMethod(null, null, null);
            dataItemValueProvider.Received(1).GetValue("Field", dataItem);

            dataItemValueProvider.ClearReceivedCalls();
            ExceptionAssert.Throws<InvalidOperationException>(() => templateProcessor.GetValue("{di:Field}"), "Template \"{di:Field}\" contains data reference but dataItem is null");
            ExceptionAssert.Throws<InvalidOperationException>(() => new DefaultTemplateProcessor(propertyValueProvider).GetValue("{di:Field}", dataItem), "Template \"{di:Field}\" contains data reference but dataItemValueProvider is null");
        }
    }
}