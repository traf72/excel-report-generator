namespace ReportEngine.Excel
{
    public struct CellCoords
    {
        public CellCoords(int rowNum, int colNum)
        {
            RowNum = rowNum;
            ColNum = colNum;
        }

        public int RowNum { get; set; }

        public int ColNum { get; set; }
    }
}