namespace LiteExcel;

/// <summary>
/// 单元格区域（用于合并单元格等）
/// 行列号均为 0-based
/// </summary>
public sealed class CellRange
{
    public int FirstRow { get; set; }
    public int LastRow { get; set; }
    public int FirstCol { get; set; }
    public int LastCol { get; set; }

    public CellRange() { }

    public CellRange(int firstRow, int lastRow, int firstCol, int lastCol)
    {
        FirstRow = firstRow;
        LastRow = lastRow;
        FirstCol = firstCol;
        LastCol = lastCol;
    }
}
