using ExcelReporter.Enums;
using System;

namespace ExcelReporter.Rendering
{
    /// <summary>
    /// Describes excel dynamic column
    /// </summary>
    internal struct ExcelDynamicColumn
    {
        private string _caption;

        public ExcelDynamicColumn(string name, Type dataType = null, string caption = null)
        {
            Name = name;
            DataType = dataType;
            _caption = caption;
            Width = null;
            AggregateFunction = dataType == typeof(decimal) || dataType == typeof(decimal?)
                ? AggregateFunction.Sum
                : AggregateFunction.NoAggregation;
            Order = 0;
        }

        /// <summary>
        /// Column name from data source
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Column caption which will be displayed in excel sheet
        /// </summary>
        public string Caption
        {
            get => _caption ?? Name;
            set => _caption = value;
        }

        /// <summary>
        /// Column width
        /// </summary>
        public double? Width { get; set; }

        /// <summary>
        /// Column data type
        /// </summary>
        public Type DataType { get; set; }

        /// <summary>
        /// Aggregate function applied to this column
        /// </summary>
        public AggregateFunction AggregateFunction { get; set; }

        /// <summary>
        /// Order in which the column appears in Excel
        /// </summary>
        public int Order { get; set; }
    }
}