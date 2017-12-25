using ExcelReporter.Enums;
using System;

namespace ExcelReporter.Attributes
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    public class ExcelColumnAttribute : Attribute
    {
        public ExcelColumnAttribute()
        {
            AggregateFunction = AggregateFunction.NoAggregation;
        }

        /// <summary>
        /// Column caption which will be shown in excel
        /// </summary>
        public string Caption { get; set; }

        /// <summary>
        /// Column width
        /// </summary>
        public double Width { get; set; }

        /// <summary>
        /// Aggregate function applied to this column
        /// </summary>
        public AggregateFunction AggregateFunction { get; set; }

        /// <summary>
        /// Do not apply an aggregate function even if it is specified
        /// </summary>
        public bool NoAggregate { get; set; }

        /// <summary>
        /// Order in which the column appears in Excel 
        /// </summary>
        public int Order { get; set; }
    }
}