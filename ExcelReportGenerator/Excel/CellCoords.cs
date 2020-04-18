namespace ExcelReportGenerator.Excel
{
    internal struct CellCoords
    {
        public CellCoords(int rowNum, int colNum)
        {
            RowNum = rowNum;
            ColNum = colNum;
        }

        public int RowNum { get; }

        public int ColNum { get; }
    }
}