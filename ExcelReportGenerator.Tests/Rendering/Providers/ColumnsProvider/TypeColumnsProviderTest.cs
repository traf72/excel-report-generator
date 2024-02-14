using ExcelReportGenerator.Attributes;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Rendering.Providers.ColumnsProviders;
using Assert = NUnit.Framework.Legacy.ClassicAssert;

namespace ExcelReportGenerator.Tests.Rendering.Providers.ColumnsProvider;

public class TypeColumnsProviderTest
{
    [Test]
    public void TestGetColumnsList()
    {
        IColumnsProvider columnsProvider = new TypeColumnsProvider();
        var columns = columnsProvider.GetColumnsList(typeof(TestType));

        Assert.AreEqual(11, columns.Count);

        Assert.AreEqual("Column4", columns[0].Name);
        Assert.AreEqual("Column4", columns[0].Caption);
        Assert.AreEqual(typeof(decimal), columns[0].DataType);
        Assert.AreEqual(AggregateFunction.Sum, columns[0].AggregateFunction);
        Assert.IsNull(columns[0].Width);
        Assert.AreEqual("#,0.00", columns[0].DisplayFormat);
        Assert.IsFalse(columns[0].AdjustToContent);
        Assert.AreEqual(0, columns[0].Order);

        Assert.AreEqual("Column5", columns[1].Name);
        Assert.AreEqual("Column5", columns[1].Caption);
        Assert.AreEqual(typeof(decimal?), columns[1].DataType);
        Assert.AreEqual(AggregateFunction.Sum, columns[1].AggregateFunction);
        Assert.IsNull(columns[1].Width);
        Assert.AreEqual("#,0.00", columns[1].DisplayFormat);
        Assert.IsFalse(columns[1].AdjustToContent);
        Assert.AreEqual(0, columns[1].Order);

        Assert.AreEqual("OverriddenColumn2", columns[2].Name);
        Assert.AreEqual("OverriddenColumn2", columns[2].Caption);
        Assert.AreEqual(typeof(string), columns[2].DataType);
        Assert.AreEqual(AggregateFunction.NoAggregation, columns[2].AggregateFunction);
        Assert.IsNull(columns[2].Width);
        Assert.IsNull(columns[2].DisplayFormat);
        Assert.IsFalse(columns[2].AdjustToContent);
        Assert.AreEqual(0, columns[2].Order);

        Assert.AreEqual("Column2", columns[3].Name);
        Assert.AreEqual("Column Two", columns[3].Caption);
        Assert.AreEqual(typeof(int), columns[3].DataType);
        Assert.AreEqual(AggregateFunction.Count, columns[3].AggregateFunction);
        Assert.IsNull(columns[3].Width);
        Assert.AreEqual("0", columns[3].DisplayFormat);
        Assert.IsFalse(columns[3].AdjustToContent);
        Assert.AreEqual(1, columns[3].Order);

        Assert.AreEqual("Column1", columns[4].Name);
        Assert.AreEqual("Column1", columns[4].Caption);
        Assert.AreEqual(typeof(string), columns[4].DataType);
        Assert.AreEqual(AggregateFunction.NoAggregation, columns[4].AggregateFunction);
        Assert.IsNull(columns[4].Width);
        Assert.IsNull(columns[4].DisplayFormat);
        Assert.IsFalse(columns[4].AdjustToContent);
        Assert.AreEqual(2, columns[4].Order);

        Assert.AreEqual("Column3", columns[5].Name);
        Assert.AreEqual("Column Three", columns[5].Caption);
        Assert.AreEqual(typeof(decimal?), columns[5].DataType);
        Assert.AreEqual(AggregateFunction.NoAggregation, columns[5].AggregateFunction);
        Assert.AreEqual(100.5, columns[5].Width);
        Assert.IsNull(columns[5].DisplayFormat);
        Assert.IsTrue(columns[5].AdjustToContent);
        Assert.AreEqual(3, columns[5].Order);

        Assert.AreEqual("OverriddenColumn3", columns[6].Name);
        Assert.AreEqual("OverriddenColumn3", columns[6].Caption);
        Assert.AreEqual(typeof(string), columns[6].DataType);
        Assert.AreEqual(AggregateFunction.Max, columns[6].AggregateFunction);
        Assert.IsNull(columns[6].Width);
        Assert.IsNull(columns[6].DisplayFormat);
        Assert.IsFalse(columns[6].AdjustToContent);
        Assert.AreEqual(4, columns[6].Order);

        Assert.AreEqual("ColumnWithBadWidth", columns[7].Name);
        Assert.AreEqual("ColumnWithBadWidth", columns[7].Caption);
        Assert.AreEqual(typeof(short), columns[7].DataType);
        Assert.AreEqual(AggregateFunction.NoAggregation, columns[7].AggregateFunction);
        Assert.IsNull(columns[7].Width);
        Assert.IsNull(columns[7].DisplayFormat);
        Assert.IsFalse(columns[7].AdjustToContent);
        Assert.AreEqual(6, columns[7].Order);

        Assert.AreEqual("OverriddenColumn", columns[8].Name);
        Assert.AreEqual("OverriddenColumn", columns[8].Caption);
        Assert.AreEqual(typeof(string), columns[8].DataType);
        Assert.AreEqual(AggregateFunction.NoAggregation, columns[8].AggregateFunction);
        Assert.IsNull(columns[8].Width);
        Assert.IsNull(columns[8].DisplayFormat);
        Assert.IsFalse(columns[8].AdjustToContent);
        Assert.AreEqual(7, columns[8].Order);

        Assert.AreEqual("Column6", columns[9].Name);
        Assert.AreEqual("Column6", columns[9].Caption);
        Assert.AreEqual(typeof(decimal), columns[9].DataType);
        Assert.AreEqual(AggregateFunction.Avg, columns[9].AggregateFunction);
        Assert.IsNull(columns[9].Width);
        Assert.AreEqual("#,##", columns[9].DisplayFormat);
        Assert.IsFalse(columns[9].AdjustToContent);
        Assert.AreEqual(9, columns[9].Order);

        Assert.AreEqual("OverriddenColumn1", columns[10].Name);
        Assert.AreEqual("OverriddenColumn1", columns[10].Caption);
        Assert.AreEqual(typeof(string), columns[10].DataType);
        Assert.AreEqual(AggregateFunction.NoAggregation, columns[10].AggregateFunction);
        Assert.IsNull(columns[10].Width);
        Assert.IsNull(columns[10].DisplayFormat);
        Assert.IsFalse(columns[10].AdjustToContent);
        Assert.AreEqual(100, columns[10].Order);
    }

    [Test]
    public void TestGetColumnsListIfTypeIsNull()
    {
        IColumnsProvider columnsProvider = new TypeColumnsProvider();
        Assert.AreEqual(0, columnsProvider.GetColumnsList(null).Count);
    }

    internal class TestType : TestTypeBase
    {
        [ExcelColumn(Order = 2)] public string Column1 = null;

        public decimal Column4 = 0;

        [ExcelColumn(Order = 1, Caption = "Column Two", AggregateFunction = AggregateFunction.Count,
            DisplayFormat = "0")]
        public int Column2 { get; set; }

        [ExcelColumn(Order = 3, Caption = "Column Three", Width = 100.5, AdjustToContent = true, NoAggregate = true,
            IgnoreDisplayFormat = true)]
        public decimal? Column3 { get; set; }

        public decimal? Column5 { get; set; }

        [ExcelColumn(Order = 9, AggregateFunction = AggregateFunction.Avg, DisplayFormat = "#,##")]
        public decimal Column6 { get; set; }

        public override string OverriddenColumn { get; set; }

        [ExcelColumn(Order = 100)] public override string OverriddenColumn1 { get; set; }

        public override string OverriddenColumn2 { get; set; }

        [ExcelColumn(Order = 4, AggregateFunction = AggregateFunction.Max)]
        public override string OverriddenColumn3 { get; set; }

        [NoExcelColumn] public string NotColumn { get; set; }

        [ExcelColumn(Order = 5)] private string NotColumn2 { get; set; }
    }

    internal class TestTypeBase
    {
        [ExcelColumn(Order = 6, Width = -10, AggregateFunction = AggregateFunction.Avg, NoAggregate = true)]
        public short ColumnWithBadWidth = 0;

        [ExcelColumn(Order = 7)] public virtual string OverriddenColumn { get; set; }

        [ExcelColumn(Order = 8)] public virtual string OverriddenColumn1 { get; set; }

        [NoExcelColumn] public virtual string OverriddenColumn2 { get; set; }

        [NoExcelColumn] public virtual string OverriddenColumn3 { get; set; }
    }
}