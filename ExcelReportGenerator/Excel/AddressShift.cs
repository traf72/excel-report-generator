namespace ExcelReportGenerator.Excel;

internal struct AddressShift
{
    public AddressShift(int rowCount, int colCount)
    {
        RowCount = rowCount;
        ColCount = colCount;
    }

    public int RowCount { get; }

    public int ColCount { get; }
}