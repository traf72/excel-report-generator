using ExcelReporter.Attributes;
using ExcelReporter.Enums;
using ExcelReporter.Rendering;
using ExcelReporter.Rendering.Providers.ColumnsProviders;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace ExcelReporter.Tests.Rendering.Providers.ColumnsProvider
{
    [TestClass]
    public class TypeColumnsProviderTest
    {
        [TestMethod]
        public void TestGetColumnsList()
        {
            IColumnsProvider columnsProvider = new TypeColumnsProvider();
            IList<ExcelDynamicColumn> columns = columnsProvider.GetColumnsList(typeof(TestType));

            Assert.AreEqual(10, columns.Count);

            Assert.AreEqual("Column4", columns[0].Name);
            Assert.AreEqual("Column4", columns[0].Caption);
            Assert.AreEqual(typeof(decimal), columns[0].DataType);
            Assert.AreEqual(AggregateFunction.Sum, columns[0].AggregateFunction);
            Assert.IsNull(columns[0].Width);
            Assert.AreEqual(0, columns[0].Order);

            Assert.AreEqual("Column5", columns[1].Name);
            Assert.AreEqual("Column5", columns[1].Caption);
            Assert.AreEqual(typeof(decimal?), columns[1].DataType);
            Assert.AreEqual(AggregateFunction.Sum, columns[1].AggregateFunction);
            Assert.IsNull(columns[1].Width);
            Assert.AreEqual(0, columns[1].Order);

            Assert.AreEqual("OverriddenColumn2", columns[2].Name);
            Assert.AreEqual("OverriddenColumn2", columns[2].Caption);
            Assert.AreEqual(typeof(string), columns[2].DataType);
            Assert.AreEqual(AggregateFunction.NoAggregation, columns[2].AggregateFunction);
            Assert.IsNull(columns[2].Width);
            Assert.AreEqual(0, columns[2].Order);

            Assert.AreEqual("Column2", columns[3].Name);
            Assert.AreEqual("Column Two", columns[3].Caption);
            Assert.AreEqual(typeof(int), columns[3].DataType);
            Assert.AreEqual(AggregateFunction.Count, columns[3].AggregateFunction);
            Assert.IsNull(columns[3].Width);
            Assert.AreEqual(1, columns[3].Order);

            Assert.AreEqual("Column1", columns[4].Name);
            Assert.AreEqual("Column1", columns[4].Caption);
            Assert.AreEqual(typeof(string), columns[4].DataType);
            Assert.AreEqual(AggregateFunction.NoAggregation, columns[4].AggregateFunction);
            Assert.IsNull(columns[4].Width);
            Assert.AreEqual(2, columns[4].Order);

            Assert.AreEqual("Column3", columns[5].Name);
            Assert.AreEqual("Column Three", columns[5].Caption);
            Assert.AreEqual(typeof(decimal?), columns[5].DataType);
            Assert.AreEqual(AggregateFunction.NoAggregation, columns[5].AggregateFunction);
            Assert.AreEqual(100.5, columns[5].Width);
            Assert.AreEqual(3, columns[5].Order);

            Assert.AreEqual("OverriddenColumn3", columns[6].Name);
            Assert.AreEqual("OverriddenColumn3", columns[6].Caption);
            Assert.AreEqual(typeof(string), columns[6].DataType);
            Assert.AreEqual(AggregateFunction.Max, columns[6].AggregateFunction);
            Assert.IsNull(columns[6].Width);
            Assert.AreEqual(4, columns[6].Order);

            Assert.AreEqual("ColumnWithBadWidth", columns[7].Name);
            Assert.AreEqual("ColumnWithBadWidth", columns[7].Caption);
            Assert.AreEqual(typeof(short), columns[7].DataType);
            Assert.AreEqual(AggregateFunction.NoAggregation, columns[7].AggregateFunction);
            Assert.IsNull(columns[7].Width);
            Assert.AreEqual(6, columns[7].Order);

            Assert.AreEqual("OverriddenColumn", columns[8].Name);
            Assert.AreEqual("OverriddenColumn", columns[8].Caption);
            Assert.AreEqual(typeof(string), columns[8].DataType);
            Assert.AreEqual(AggregateFunction.NoAggregation, columns[8].AggregateFunction);
            Assert.IsNull(columns[8].Width);
            Assert.AreEqual(7, columns[8].Order);

            Assert.AreEqual("OverriddenColumn1", columns[9].Name);
            Assert.AreEqual("OverriddenColumn1", columns[9].Caption);
            Assert.AreEqual(typeof(string), columns[9].DataType);
            Assert.AreEqual(AggregateFunction.NoAggregation, columns[9].AggregateFunction);
            Assert.IsNull(columns[9].Width);
            Assert.AreEqual(100, columns[9].Order);
        }

        [TestMethod]
        public void TestGetColumnsListIfTypeIsNull()
        {
            IColumnsProvider columnsProvider = new TypeColumnsProvider();
            Assert.AreEqual(0, columnsProvider.GetColumnsList(null).Count);
        }

        internal class TestType : TestTypeBase
        {
            [ExcelColumn(Order = 2)]
            public string Column1 = null;

            [ExcelColumn(Order = 1, Caption = "Column Two", AggregateFunction = AggregateFunction.Count)]
            public int Column2 { get; set; }

            [ExcelColumn(Order = 3, Caption = "Column Three", Width = 100.5, NoAggregate = true)]
            public decimal? Column3 { get; set; }

            public decimal Column4 = 0;

            public decimal? Column5 { get; set; }

            public override string OverriddenColumn { get; set; }

            [ExcelColumn(Order = 100)]
            public override string OverriddenColumn1 { get; set; }

            public override string OverriddenColumn2 { get; set; }

            [ExcelColumn(Order = 4, AggregateFunction = AggregateFunction.Max)]
            public override string OverriddenColumn3 { get; set; }

            [NoExcelColumn]
            public string NotColumn { get; set; }

            [ExcelColumn(Order = 5)]
            private string NotColumn2 { get; set; }
        }

        internal class TestTypeBase
        {
            [ExcelColumn(Order = 6, Width = -10, AggregateFunction = AggregateFunction.Avg, NoAggregate = true)]
            public short ColumnWithBadWidth = 0;

            [ExcelColumn(Order = 7)]
            public virtual string OverriddenColumn { get; set; }

            [ExcelColumn(Order = 8)]
            public virtual string OverriddenColumn1 { get; set; }

            [NoExcelColumn]
            public virtual string OverriddenColumn2 { get; set; }

            [NoExcelColumn]
            public virtual string OverriddenColumn3 { get; set; }
        }
    }
}