﻿using System.Data;
using ExcelReportGenerator.Exceptions;
using ExcelReportGenerator.Rendering;
using ExcelReportGenerator.Rendering.Providers.DataItemValueProviders;
using ExcelReportGenerator.Tests.CustomAsserts;

namespace ExcelReportGenerator.Tests.Rendering.Providers.DataItemValueProviders;

public class HierarchicalDataItemValueProviderTest
{
    [Test]
    public void TestGetValue()
    {
        var dataTable = new DataTable();
        dataTable.Columns.Add("Column", typeof(int));
        dataTable.Rows.Add(1);

        var dataItem1 = new TestClass2();
        var dataItem2 = dataTable.Rows[0];
        var dataItem3 = Substitute.For<IDataReader>();
        var dataItem4 = new TestClass();

        var hierarchicalDataItem = new HierarchicalDataItem
        {
            Value = dataItem4,
            Parent = new HierarchicalDataItem
            {
                Value = dataItem1,
                Parent = new HierarchicalDataItem
                {
                    Value = dataItem2,
                    Parent = new HierarchicalDataItem
                    {
                        Value = dataItem3
                    }
                }
            }
        };

        var factory = Substitute.For<IDataItemValueProviderFactory>();
        var objectPropertyValueProvider = Substitute.For<ObjectPropertyValueProvider>();
        var dataRowValueProvider = Substitute.For<DataRowValueProvider>();
        var dataReaderValueProvider = Substitute.For<DataReaderValueProvider>();
        var dataItemValueProvider = new DefaultDataItemValueProvider(factory) {DataItemSelfTemplate = "di"};

        factory.Create(null).Returns(objectPropertyValueProvider);
        factory.Create(dataItem1).Returns(objectPropertyValueProvider);
        factory.Create(dataItem2).Returns(dataRowValueProvider);
        factory.Create(dataItem3).Returns(dataReaderValueProvider);
        factory.Create(dataItem4).Returns(objectPropertyValueProvider);

        ExceptionAssert.Throws<ArgumentException>(() => dataItemValueProvider.GetValue(null, hierarchicalDataItem));
        ExceptionAssert.Throws<ArgumentException>(() =>
            dataItemValueProvider.GetValue(string.Empty, hierarchicalDataItem));
        ExceptionAssert.Throws<ArgumentException>(() => dataItemValueProvider.GetValue(" ", hierarchicalDataItem));

        ExceptionAssert.Throws<ArgumentNullException>(() => dataItemValueProvider.GetValue("Template", null));
        ExceptionAssert.Throws<ArgumentNullException>(() => dataItemValueProvider.GetValue("Template", null));
        ExceptionAssert.Throws<ArgumentNullException>(() => dataItemValueProvider.GetValue("Template", null));

        objectPropertyValueProvider.ClearReceivedCalls();
        dataItemValueProvider.GetValue("di", hierarchicalDataItem);
        objectPropertyValueProvider.DidNotReceiveWithAnyArgs().GetValue(null, null);
        dataRowValueProvider.DidNotReceiveWithAnyArgs().GetValue(null, null);
        dataReaderValueProvider.DidNotReceiveWithAnyArgs().GetValue(null, null);

        dataItemValueProvider.GetValue("Prop", hierarchicalDataItem);
        objectPropertyValueProvider.Received(1).GetValue("Prop", dataItem4);
        dataRowValueProvider.DidNotReceiveWithAnyArgs().GetValue(null, null);
        dataReaderValueProvider.DidNotReceiveWithAnyArgs().GetValue(null, null);

        dataItemValueProvider.DataItemSelfTemplate = "dataItem";
        objectPropertyValueProvider.ClearReceivedCalls();
        dataItemValueProvider.GetValue("parent:dataItem", hierarchicalDataItem);
        objectPropertyValueProvider.DidNotReceiveWithAnyArgs().GetValue(null, null);
        dataRowValueProvider.DidNotReceiveWithAnyArgs().GetValue(null, null);
        dataReaderValueProvider.DidNotReceiveWithAnyArgs().GetValue(null, null);

        dataItemValueProvider.GetValue("parent:Prop", hierarchicalDataItem);
        objectPropertyValueProvider.Received(1).GetValue("Prop", dataItem1);
        dataRowValueProvider.DidNotReceiveWithAnyArgs().GetValue(null, null);
        dataReaderValueProvider.DidNotReceiveWithAnyArgs().GetValue(null, null);

        objectPropertyValueProvider.ClearReceivedCalls();
        dataItemValueProvider.GetValue("parent:parent", hierarchicalDataItem);
        objectPropertyValueProvider.Received(1).GetValue("parent", dataItem1);
        dataRowValueProvider.DidNotReceiveWithAnyArgs().GetValue(null, null);
        dataReaderValueProvider.DidNotReceiveWithAnyArgs().GetValue(null, null);

        objectPropertyValueProvider.ClearReceivedCalls();
        dataItemValueProvider.GetValue("parent : PARENT: Column", hierarchicalDataItem);
        objectPropertyValueProvider.DidNotReceiveWithAnyArgs().GetValue(null, null);
        dataRowValueProvider.Received(1).GetValue("Column", dataItem2);
        dataReaderValueProvider.DidNotReceiveWithAnyArgs().GetValue(null, null);

        dataRowValueProvider.ClearReceivedCalls();
        dataItemValueProvider.GetValue("parent : PARENT :parent:Column", hierarchicalDataItem);
        objectPropertyValueProvider.DidNotReceiveWithAnyArgs().GetValue(null, null);
        dataRowValueProvider.DidNotReceiveWithAnyArgs().GetValue(null, null);
        dataReaderValueProvider.Received(1).GetValue("Column", dataItem3);

        hierarchicalDataItem.Value = null;
        dataReaderValueProvider.ClearReceivedCalls();
        dataItemValueProvider.GetValue("Prop", hierarchicalDataItem);
        objectPropertyValueProvider.Received(1).GetValue("Prop", null);
        dataRowValueProvider.DidNotReceiveWithAnyArgs().GetValue(null, null);
        dataReaderValueProvider.DidNotReceiveWithAnyArgs().GetValue(null, null);

        ExceptionAssert.Throws<InvalidTemplateException>(
            () => dataItemValueProvider.GetValue("par:Prop", hierarchicalDataItem), "Template \"par:Prop\" is invalid");
    }

    private class TestClass
    {
    }

    private class TestClass2
    {
    }
}