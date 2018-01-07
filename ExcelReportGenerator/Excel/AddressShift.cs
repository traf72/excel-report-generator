﻿namespace ExcelReportGenerator.Excel
{
    internal struct AddressShift
    {
        public AddressShift(int rowCount, int colCount)
        {
            RowCount = rowCount;
            ColCount = colCount;
        }

        public int RowCount { get; set; }

        public int ColCount { get; set; }
    }
}