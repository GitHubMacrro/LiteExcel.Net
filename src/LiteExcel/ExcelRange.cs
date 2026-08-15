using System.Collections;

namespace LiteExcel;

/// <summary>
/// 连续矩形区域（1-based，含端点）。
/// 支持批量读写、样式、合并、清空、枚举。
/// </summary>
public sealed class ExcelRange : IEnumerable<Cell>
{
    private readonly Worksheet _sheet;

    /// <summary>区域首行（1-based） </summary>
    public int FirstRow { get; }

    /// <summary>区域首列（1-based） </summary>
    public int FirstCol { get; }

    /// <summary>区域末行（1-based） </summary>
    public int LastRow { get; }

    /// <summary>区域末列（1-based） </summary>
    public int LastCol { get; }

    internal ExcelRange(Worksheet sheet, int firstRow, int firstCol, int lastRow, int lastCol)
    {
        _sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
        FirstRow = firstRow;
        FirstCol = firstCol;
        LastRow = lastRow;
        LastCol = lastCol;
    }

    /// <summary>区域 A1 地址，如 "A1:D10" </summary>
    public string Address => $"{CellRef.ToString(FirstRow - 1, FirstCol - 1)}:{CellRef.ToString(LastRow - 1, LastCol - 1)}";

    /// <summary>行数 </summary>
    public int RowCount => LastRow - FirstRow + 1;

    /// <summary>列数 </summary>
    public int ColumnCount => LastCol - FirstCol + 1;

    /// <summary>区域内的单元格（相对偏移，0-based） </summary>
    public Cell Cell(int rowOffset, int colOffset)
    {
        if (rowOffset < 0 || rowOffset >= RowCount || colOffset < 0 || colOffset >= ColumnCount)
            throw new ArgumentOutOfRangeException(nameof(rowOffset), $"偏移超出区域：({rowOffset},{colOffset})");
        return _sheet.Cell(FirstRow + rowOffset, FirstCol + colOffset);
    }

    /// <summary>批量写入相同值（越界自动扩展） </summary>
    public void Fill(object? value)
    {
        for (int r = FirstRow; r <= LastRow; r++)
            for (int c = FirstCol; c <= LastCol; c++)
                _sheet.SetValue(r, c, value);
    }

    /// <summary>把二维数据写入区域。data[r][c] 对应区域内第 r 行第 c 列（0-based） </summary>
    public void Fill(object?[,] data)
    {
        if (data is null) throw new ArgumentNullException(nameof(data));
        for (int r = 0; r < data.GetLength(0) && r < RowCount; r++)
            for (int c = 0; c < data.GetLength(1) && c < ColumnCount; c++)
                _sheet.SetValue(FirstRow + r, FirstCol + c, data[r, c]);
    }

    /// <summary>读取区域内所有值（object?[,]），尺寸 = RowCount × ColumnCount </summary>
    public object?[,] ToValues()
    {
        var values = new object?[RowCount, ColumnCount];
        for (int r = 0; r < RowCount; r++)
            for (int c = 0; c < ColumnCount; c++)
                values[r, c] = _sheet.Cell(FirstRow + r, FirstCol + c).GetValue();
        return values;
    }

    /// <summary>读取区域内所有 Cell </summary>
    public Cell[,] ToCells()
    {
        var cells = new Cell[RowCount, ColumnCount];
        for (int r = 0; r < RowCount; r++)
            for (int c = 0; c < ColumnCount; c++)
                cells[r, c] = _sheet.Cell(FirstRow + r, FirstCol + c);
        return cells;
    }

    /// <summary>清空区域内所有单元格的值 </summary>
    public void Clear()
    {
        for (int r = FirstRow; r <= LastRow; r++)
            for (int c = FirstCol; c <= LastCol; c++)
                _sheet.Cell(r, c).SetValue(null);
    }

    /// <summary>为区域内所有单元格应用统一样式 </summary>
    public CellStyle? Style
    {
        get => _sheet.Cell(FirstRow, FirstCol).Style;
        set
        {
            for (int r = FirstRow; r <= LastRow; r++)
                for (int c = FirstCol; c <= LastCol; c++)
                    _sheet.Cell(r, c).Style = value;
        }
    }

    /// <summary>合并该区域 </summary>
    public void Merge() => _sheet.Merge(FirstRow, FirstCol, LastRow, LastCol);

    /// <summary>取消该区域合并 </summary>
    public void Unmerge() => _sheet.Unmerge(Address);

    /// <summary>枚举区域内所有单元格（按行优先） </summary>
    public IEnumerator<Cell> GetEnumerator()
    {
        for (int r = FirstRow; r <= LastRow; r++)
            for (int c = FirstCol; c <= LastCol; c++)
                yield return _sheet.Cell(r, c);
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
