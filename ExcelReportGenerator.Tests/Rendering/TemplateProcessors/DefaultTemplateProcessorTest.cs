using System;
using ExcelReportGenerator.Exceptions;
using ExcelReportGenerator.Rendering;
using ExcelReportGenerator.Rendering.Providers;
using ExcelReportGenerator.Rendering.Providers.DataItemValueProviders;
using ExcelReportGenerator.Rendering.Providers.VariableProviders;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using ExcelReportGenerator.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;

namespace ExcelReportGenerator.Tests.Rendering.TemplateProcessors
{
    [TestClass]
    public class DefaultTemplateProcessorTest
    {
        [TestMethod]
        public void TestGetValue()
        {
            var propertyValueProvider = Substitute.For<IPropertyValueProvider>();
            var variableValueProvider = Substitute.For<SystemVariableProvider>();
            var methodCallValueProvider = Substitute.For<IMethodCallValueProvider>();
            var dataItemValueProvider = Substitute.For<IGenericDataItemValueProvider<HierarchicalDataItem>>();
            var dataItem = new HierarchicalDataItem();

            ExceptionAssert.Throws<ArgumentNullException>(() => new DefaultTemplateProcessor(null, variableValueProvider, methodCallValueProvider, dataItemValueProvider));
            ExceptionAssert.Throws<ArgumentNullException>(() => new DefaultTemplateProcessor(propertyValueProvider, null, methodCallValueProvider, dataItemValueProvider));
            new DefaultTemplateProcessor(propertyValueProvider, variableValueProvider, null, dataItemValueProvider);
            new DefaultTemplateProcessor(propertyValueProvider, variableValueProvider, methodCallValueProvider);

            var templateProcessor = new DefaultTemplateProcessor(propertyValueProvider, variableValueProvider, methodCallValueProvider, dataItemValueProvider);

            ExceptionAssert.Throws<ArgumentNullException>(() => templateProcessor.GetValue(null, dataItem));
            ExceptionAssert.Throws<InvalidTemplateException>(() => templateProcessor.GetValue("{p-Name}"), "Invalid template \"{p-Name}\". Cannot find separator \":\" between member label and member template");
            ExceptionAssert.Throws<InvalidTemplateException>(() => templateProcessor.GetValue("{bad:Name}"), "Invalid template \"{bad:Name}\". Unknown member label \"bad\"");

            templateProcessor.GetValue("{p:Name}");

            propertyValueProvider.Received(1).GetValue("Name");
            variableValueProvider.DidNotReceiveWithAnyArgs().GetVariable(Arg.Any<string>());
            methodCallValueProvider.DidNotReceiveWithAnyArgs().CallMethod(null, null, null);
            dataItemValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>(), null);

            propertyValueProvider.ClearReceivedCalls();
            templateProcessor.GetValue(" { p : Name } ");

            propertyValueProvider.Received(1).GetValue("Name");
            variableValueProvider.DidNotReceiveWithAnyArgs().GetVariable(Arg.Any<string>());
            methodCallValueProvider.DidNotReceiveWithAnyArgs().CallMethod(null, null, null);
            dataItemValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>(), null);

            propertyValueProvider.ClearReceivedCalls();
            templateProcessor.GetValue(" p:Name ");

            propertyValueProvider.Received(1).GetValue("Name");
            variableValueProvider.DidNotReceiveWithAnyArgs().GetVariable(Arg.Any<string>());
            methodCallValueProvider.DidNotReceiveWithAnyArgs().CallMethod(null, null, null);
            dataItemValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>(), null);

            propertyValueProvider.ClearReceivedCalls();
            templateProcessor.GetValue("sv:Now");

            variableValueProvider.Received(1).GetVariable("Now");
            propertyValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());
            methodCallValueProvider.DidNotReceiveWithAnyArgs().CallMethod(null, null, null);
            dataItemValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>(), null);

            variableValueProvider.ClearReceivedCalls();
            templateProcessor.GetValue("sf:Format(5567.88, 0)");

            variableValueProvider.DidNotReceiveWithAnyArgs().GetVariable(Arg.Any<string>());
            propertyValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());
            methodCallValueProvider.Received(1).CallMethod("Format(5567.88, 0)", typeof(SystemFunctions), templateProcessor, null);
            dataItemValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>(), null);

            methodCallValueProvider.ClearReceivedCalls();
            templateProcessor.GetValue("{ m:Method() }", dataItem);

            propertyValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());
            variableValueProvider.DidNotReceiveWithAnyArgs().GetVariable(Arg.Any<string>());
            methodCallValueProvider.Received(1).CallMethod("Method()", templateProcessor, dataItem);
            dataItemValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>(), null);

            methodCallValueProvider.ClearReceivedCalls();
            templateProcessor.GetValue("{m:Method()}", dataItem);
            propertyValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());
            variableValueProvider.DidNotReceiveWithAnyArgs().GetVariable(Arg.Any<string>());
            methodCallValueProvider.Received(1).CallMethod("Method()", templateProcessor, dataItem);
            dataItemValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>(), null);

            methodCallValueProvider.ClearReceivedCalls();
            ExceptionAssert.Throws<InvalidOperationException>(() => new DefaultTemplateProcessor(propertyValueProvider, variableValueProvider).GetValue("{m:Method()}"), "Template \"{m:Method()}\" contains method call but methodCallValueProvider is null");
            ExceptionAssert.Throws<InvalidOperationException>(() => new DefaultTemplateProcessor(propertyValueProvider, variableValueProvider).GetValue("{sf:Format(p:Prop, 0))}"), "Template \"{sf:Format(p:Prop, 0))}\" contains system function call but methodCallValueProvider is null");

            templateProcessor.GetValue("{di:Field}", dataItem);
            propertyValueProvider.DidNotReceiveWithAnyArgs().GetValue(Arg.Any<string>());
            variableValueProvider.DidNotReceiveWithAnyArgs().GetVariable(Arg.Any<string>());
            methodCallValueProvider.DidNotReceiveWithAnyArgs().CallMethod(null, null, null);
            dataItemValueProvider.Received(1).GetValue("Field", dataItem);

            dataItemValueProvider.ClearReceivedCalls();
            ExceptionAssert.Throws<InvalidOperationException>(() => templateProcessor.GetValue("{di:Field}"), "Template \"{di:Field}\" contains data reference but dataItem is null");
            ExceptionAssert.Throws<InvalidOperationException>(() => new DefaultTemplateProcessor(propertyValueProvider, variableValueProvider).GetValue("{di:Field}", dataItem), "Template \"{di:Field}\" contains data reference but dataItemValueProvider is null");
        }
    }
}