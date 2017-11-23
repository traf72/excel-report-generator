using ClosedXML.Excel;

namespace ExcelReporter.Interfaces.Providers.DataItemColumnsProvider
{
    internal struct ExcelDynamicColumn
    {
        private string _caption;

        public ExcelDynamicColumn(string name, string caption = null)
        {
            Name = name;
            _caption = caption;
            Width = null;
            DataType = XLCellValues.Text;
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

        // TODO эта колонка под большим вопросом
        /// <summary>
        /// Column data type
        /// </summary>
        XLCellValues DataType { get; set; }
    }
}