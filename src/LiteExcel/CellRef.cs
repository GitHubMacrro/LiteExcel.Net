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
