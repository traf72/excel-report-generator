namespace ExcelReporter.Excel
{
    public struct RangeCoords
    {
        public RangeCoords(CellCoords firstCell, CellCoords lastCell)
        {
            FirstCell = firstCell;
            LastCell = lastCell;
        }

        public CellCoords FirstCell { get; set; }

        public CellCoords LastCell { get; set; }
    }
}