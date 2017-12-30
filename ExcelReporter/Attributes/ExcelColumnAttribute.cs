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
        /// Column caption which will be displayed in excel sheet
        /// </summary>
        public string Caption { get; set; }

        /// <summary>
        /// Column width if panel is vertical or row height if panel is horizontal
        /// </summary>
        public double Width { get; set; }

        /// <summary>
        /// Aggregate function applied to this column
        /// </summary>
        public AggregateFunction AggregateFunction { get; set; }

        /// <summary>
        /// Do not apply an aggregate function even it is specified
        /// </summary>
        public bool NoAggregate { get; set; }

        /// <summary>
        /// Display format for number and date columns
        /// </summary>
        public string DisplayFormat { get; set; }

        /// <summary>
        /// Do not apply display format even it is specified
        /// </summary>
        public bool IgnoreDisplayFormat { get; set; }

        /// <summary>
        /// Order in which the column appears in Excel 
        /// </summary>
        public int Order { get; set; }
    }
}