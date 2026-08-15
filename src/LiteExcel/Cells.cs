using System.Collections;

namespace LiteExcel;

/// <summary>
/// 整表单元格集合入口。
/// 提供按坐标/地址访问、区域提取、整表枚举与批量清空。
/// </summary>
public sealed class Cells : IEnumerable<Cell>
{
    private readonly Worksheet _sheet;

    internal Cells(Worksheet sheet)
    {
        _sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
    }

    /// <summary>按 1-based 行列访问单元格，如 cells[1, 1] == A1 </summary>
    public Cell this[int row, int column]
    {
        get => _sheet.Cell(row, column);
        set => _sheet.SetCell(row, column, value);
    }

    /// <summary>按 A1 地址访问单格，如 cells["A1"]。区域地址请用 <see cref="Range"/> </summary>
    public Cell this[string address]
    {
        get => _sheet.Cell(address);
        set => _sheet.SetValue(address, value);
    }

    /// <summary>按 A1 区域提取 ExcelRange，如 cells.Range("A1:D10") </summary>
    public ExcelRange Range(string address) => _sheet.Range(address);

    /// <summary>按 1-based 行列提取 ExcelRange（含端点） </summary>
    public ExcelRange Range(int firstRow, int firstCol, int lastRow, int lastCol) =>
        _sheet.Range(firstRow, firstCol, lastRow, lastCol);

    /// <summary>设置单元格值（越界自动扩展网格） </summary>
    public void SetValue(int row, int column, object? value) => _sheet.SetValue(row, column, value);

    /// <summary>按 A1 地址设置单元格值 </summary>
    public void SetValue(string address, object? value) => _sheet.SetValue(address, value);

    /// <summary>清空整表所有单元格（值置空，不删除行列） </summary>
    public void Clear()
    {
        var snapshot = _sheet.EnumerateStoredCells().ToList();
        foreach (var cell in snapshot)
            cell.SetValue(null);
    }

    /// <summary>枚举网格中已有的所有单元格 </summary>
    public IEnumerator<Cell> GetEnumerator() => _sheet.EnumerateStoredCells().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
