using System;
using System.Collections.Generic;
using System.Text;

namespace LiteExcel.Internal;

/// <summary>
/// 数字格式代码识别助手（xls / xlsb 共用）。
/// 将"数字单元格是否应解释为日期"的逻辑集中在此，避免两个后端各自维护一套。
/// </summary>
internal static class FormatDetector
{
    private static readonly HashSet<int> BuiltInDateFmtIds = new()
    {
        14, 15, 16, 17, 18, 19, 20, 21, 22,
        27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
        45, 46, 47,
        50, 51, 52, 53, 54, 55, 56, 57, 58,
    };

    /// <summary>
    /// 根据样式索引 + 自定义格式表生成单元格。
    /// <paramref name="xfIfmt"/> 为按 ixfe/ixStyle 索引的 numFmtId 列表（含索引为 0 的默认项）。
    /// </summary>
    public static Cell CellFromNumber(double val, int ixfe, List<int> xfIfmt,
        Dictionary<int, string> formats, bool date1904)
    {
        int ifmt = ixfe >= 0 && ixfe < xfIfmt.Count ? xfIfmt[ixfe] : 0;
        string? fmtCode = GetFormatCode(ifmt, formats);
        if (IsDateFormat(ifmt, fmtCode))
        {
            var date = date1904 ? new DateTime(1904, 1, 1).AddDays(val) : DateTime.FromOADate(val);
            return Cell.FromDate(date, fmtCode);
        }
        return Cell.FromNumber(val, fmtCode);
    }

    public static string? GetFormatCode(int ifmt, Dictionary<int, string> formats)
    {
        if (ifmt < 164) return GetBuiltInFormatCode(ifmt);
        return formats.TryGetValue(ifmt, out var code) ? code : null;
    }

    public static bool IsDateFormat(int ifmt, string? formatCode)
    {
        if (!string.IsNullOrEmpty(formatCode))
        {
            var lower = StripBrackets(formatCode!).ToLowerInvariant();
            return lower.Contains('y') || lower.Contains('d') || lower.Contains('h') || lower.Contains('s');
        }
        return BuiltInDateFmtIds.Contains(ifmt);
    }

    public static string StripBrackets(string fmt)
    {
        var sb = new StringBuilder(fmt.Length);
        bool inBracket = false;
        foreach (var ch in fmt)
        {
            if (ch == '[') { inBracket = true; continue; }
            if (ch == ']') { inBracket = false; continue; }
            if (!inBracket) sb.Append(ch);
        }
        return sb.ToString();
    }

    public static string? GetBuiltInFormatCode(int ifmt) => ifmt switch
    {
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "yyyy-MM-dd",
        15 => "dd-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "yyyy-MM-dd h:mm",
        37 => "#,##0 ;(#,##0)",
        38 => "#,##0 ;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;[Red](#,##0.00)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mmss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => null,
    };
}
