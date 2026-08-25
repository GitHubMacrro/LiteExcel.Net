using System.Text;

namespace LiteExcel;

/// <summary>
/// A1 引用与行列号（0-based）互转工具 
/// </summary>
public static class CellRef
{
    /// <summary>
    /// "A1" -> (row=0, col=0)；"B3" -> (row=2, col=1) 
    /// </summary>
    public static (int row, int col) Parse(string cellRef)
    {
        int col = 0;
        int i = 0;
        for (; i < cellRef.Length; i++)
        {
            char ch = cellRef[i];
            if (ch >= 'A' && ch <= 'Z') col = col * 26 + (ch - 'A' + 1);
            else if (ch >= 'a' && ch <= 'z') col = col * 26 + (ch - 'a' + 1);
            else break;
        }
        col--;

        int row = 0;
        for (; i < cellRef.Length; i++)
        {
            char ch = cellRef[i];
            if (ch >= '0' && ch <= '9') row = row * 10 + (ch - '0');
            else break;
        }
        row--;

        return (row, col);
    }

    /// <summary>
    /// 尝试把 A1 单元格地址解析为 (row, col)，格式非法返回 false。
    /// </summary>
    public static bool TryParse(string? cellRef, out (int row, int col) pos)
    {
        pos = default;
        if (string.IsNullOrEmpty(cellRef)) return false;
        var (row, col) = Parse(cellRef);
        if (row < 0 || col < 0) return false;
        // 重新生成校验原串是纯 A1（防 "A1C" 这类尾巴被吞）
        if (!string.Equals(ToString(row, col), cellRef, StringComparison.OrdinalIgnoreCase)) return false;
        pos = (row, col);
        return true;
    }

    /// <summary>
    /// 解析区域引用："A1" 或 "A1:D100" -> (firstRow, firstCol, lastRow, lastCol)，全 0-based 含端点。
    /// </summary>
    public static (int firstRow, int firstCol, int lastRow, int lastCol) ParseRange(string range)
    {
        if (string.IsNullOrWhiteSpace(range))
            throw new ArgumentException("区域不能为空", nameof(range));
        int colon = range.IndexOf(':');
        string first = colon >= 0 ? range.Substring(0, colon) : range;
        string last = colon >= 0 ? range.Substring(colon + 1) : range;
        var (r0, c0) = Parse(first);
        var (r1, c1) = Parse(last);
        if (r0 < 0 || c0 < 0 || r1 < 0 || c1 < 0)
            throw new ArgumentException($"无效的区域引用：{range}");
        return (Math.Min(r0, r1), Math.Min(c0, c1), Math.Max(r0, r1), Math.Max(c0, c1));
    }

    /// <summary>
    /// (row=0, col=0) -> "A1" 
    /// </summary>
    public static string ToString(int row, int col)
    {
        return ColToLetter(col) + (row + 1);
    }

    /// <summary>
    /// col=0 -> "A", col=25 -> "Z", col=26 -> "AA" 
    /// </summary>
    public static string ColToLetter(int col)
    {
        var sb = new StringBuilder();
        col++;
        while (col > 0)
        {
            col--;
            sb.Insert(0, (char)('A' + col % 26));
            col /= 26;
        }
        return sb.ToString();
    }

    /// <summary>
    /// "A" -> 0, "Z" -> 25, "AA" -> 26 
    /// </summary>
    public static int LetterToCol(string letters)
    {
        int col = 0;
        foreach (var ch in letters)
        {
            if (ch >= 'A' && ch <= 'Z') col = col * 26 + (ch - 'A' + 1);
            else if (ch >= 'a' && ch <= 'z') col = col * 26 + (ch - 'a' + 1);
            else break;
        }
        return col - 1;
    }
}
