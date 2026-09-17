using System.Text;

namespace LiteExcel.Internal.Biff12;

/// <summary>
/// BIFF12 定义名称（BrtDefinedName）的 rgce（RPN 令牌流）→ A1 引用文本解码器。
///
/// 范围（保守）：仅解码**单个**「引用 / 区域 / 常量」令牌，产出 <c>Sheet!$A$1:$B$5</c> 形式的引用文本。
/// 复合表达式（函数调用、运算符、数组、名称引用等）一律返回 null，由调用方降级上报——
/// 宁可不上报该名称，也不产出错误的引用文本（保真契约）。
///
/// 令牌布局（[MS-XLSB] + 真实 Excel 样本标定）：
///   PtgInt  0x1E : u16
///   PtgNum  0x1F : f64
///   PtgBool 0x1D : u8
///   PtgStr  0x17 : cch(u16) + UTF-16
///   PtgRef  0x44 : rw(u32) + col(u16)                      （col 位15=fRwRel, 位14=fColRel）
///   PtgArea 0x45 : rwFirst(u32) + rwLast(u32) + colFirst(u16) + colLast(u16)
///   PtgRef3d 0x3A: ixti(u16) + rw(u32) + col(u16)
///   PtgArea3d 0x3B: ixti(u16) + rwFirst(u32) + rwLast(u32) + colFirst(u16) + colLast(u16)
/// </summary>
internal static class RgceDecoder
{
    private const byte PtgStr = 0x17;
    private const byte PtgBool = 0x1D;
    private const byte PtgInt = 0x1E;
    private const byte PtgNum = 0x1F;
    private const byte PtgRef3d = 0x3A;
    private const byte PtgArea3d = 0x3B;
    private const byte PtgRef = 0x44;
    private const byte PtgArea = 0x45;

    /// <summary>
    /// 解码 rgce。成功返回 A1 引用文本；不支持的令牌序列返回 null。
    /// </summary>
    /// <param name="rgce">RPN 令牌字节</param>
    /// <param name="sheetNameByIxti">ixti → 工作表名（来自 BrtExternSheet）；ixti 越界返回 null</param>
    /// <param name="currentSheetName">当前表名（用于 3D 引用缺 sheet 时；本解码器不使用，保留签名扩展）</param>
    public static string? Decode(byte[] rgce, IReadOnlyList<string> sheetNameByIxti)
    {
        if (rgce.Length == 0) return null;
        int off = 0;
        string? result = DecodeSingle(rgce, ref off, sheetNameByIxti);
        if (result is null) return null;
        // 必须恰好消费完整个 rgce（否则是复合表达式）
        return off == rgce.Length ? result : null;
    }

    private static string? DecodeSingle(byte[] d, ref int off, IReadOnlyList<string> sheets)
    {
        if (off >= d.Length) return null;
        byte ptg = d[off];

        switch (ptg)
        {
            case PtgInt:
            {
                if (off + 3 > d.Length) return null;
                ushort v = (ushort)(d[off + 1] | (d[off + 2] << 8));
                off += 3;
                return v.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
            case PtgNum:
            {
                if (off + 9 > d.Length) return null;
                double v = BitConverter.ToDouble(d, off + 1);
                off += 9;
                return v.ToString("R", System.Globalization.CultureInfo.InvariantCulture);
            }
            case PtgBool:
            {
                if (off + 2 > d.Length) return null;
                bool b = d[off + 1] != 0;
                off += 2;
                return b ? "TRUE" : "FALSE";
            }
            case PtgStr:
            {
                if (off + 3 > d.Length) return null;
                int cch = d[off + 1] | (d[off + 2] << 8);
                if (off + 3 + cch * 2 > d.Length) return null;
                string s = Encoding.Unicode.GetString(d, off + 3, cch * 2);
                off += 3 + cch * 2;
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            }
            case PtgRef:
            {
                // rw(u32) + col(u16)
                if (off + 7 > d.Length) return null;
                uint rw = BitConverter.ToUInt32(d, off + 1);
                ushort colWord = (ushort)(d[off + 5] | (d[off + 6] << 8));
                off += 7;
                return CellRefText((int)rw, colWord);
            }
            case PtgArea:
            {
                if (off + 13 > d.Length) return null;
                uint rwF = BitConverter.ToUInt32(d, off + 1);
                uint rwL = BitConverter.ToUInt32(d, off + 5);
                ushort colF = (ushort)(d[off + 9] | (d[off + 10] << 8));
                ushort colL = (ushort)(d[off + 11] | (d[off + 12] << 8));
                off += 13;
                return CellRefText((int)rwF, colF) + ":" + CellRefText((int)rwL, colL);
            }
            case PtgRef3d:
            {
                // ixti(u16) + rw(u32) + col(u16)
                if (off + 9 > d.Length) return null;
                int ixti = d[off + 1] | (d[off + 2] << 8);
                if (ixti < 0 || ixti >= sheets.Count) return null;
                uint rw = BitConverter.ToUInt32(d, off + 3);
                ushort colWord = (ushort)(d[off + 7] | (d[off + 8] << 8));
                off += 9;
                return QuoteSheet(sheets[ixti]) + "!" + CellRefText((int)rw, colWord);
            }
            case PtgArea3d:
            {
                if (off + 15 > d.Length) return null;
                int ixti = d[off + 1] | (d[off + 2] << 8);
                if (ixti < 0 || ixti >= sheets.Count) return null;
                uint rwF = BitConverter.ToUInt32(d, off + 3);
                uint rwL = BitConverter.ToUInt32(d, off + 7);
                ushort colF = (ushort)(d[off + 11] | (d[off + 12] << 8));
                ushort colL = (ushort)(d[off + 13] | (d[off + 14] << 8));
                off += 15;
                return QuoteSheet(sheets[ixti]) + "!" + CellRefText((int)rwF, colF) + ":" + CellRefText((int)rwL, colL);
            }
            default:
                return null; // 函数/运算符/名称引用等：不支持
        }
    }

    /// <summary>BIFF12 单元格引用（rw 0-based, colWord 含相对/绝对标志）→ A1 文本。</summary>
    private static string CellRefText(int rw, ushort colWord)
    {
        bool colRel = (colWord & 0x4000) != 0;
        bool rwRel = (colWord & 0x8000) != 0;
        int col = colWord & 0x3FFF;
        string colName = ColumnName(col);
        return (colRel ? "" : "$") + colName + (rwRel ? "" : "$") + (rw + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    /// <summary>0-based 列号 → A..XFD。</summary>
    internal static string ColumnName(int col)
    {
        var buf = new char[4];
        int n = 0, c = col + 1;
        while (c > 0) { int m = (c - 1) % 26; buf[n++] = (char)('A' + m); c = (c - 1) / 26; }
        var sb = new StringBuilder(4);
        for (int i = n - 1; i >= 0; i--) sb.Append(buf[i]);
        return sb.ToString();
    }

    /// <summary>工作表名按 Excel 规则加引号（含空格/特殊字符或数字开头时）。</summary>
    private static string QuoteSheet(string name)
    {
        bool need = name.Length == 0;
        if (!need)
        {
            char first = name[0];
            if (first >= '0' && first <= '9') need = true;
            foreach (var ch in name)
            {
                bool plain = (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z')
                    || (ch >= '0' && ch <= '9') || ch == '_' || ch > 0x7F; // CJK 等视为安全
                if (!plain) { need = true; break; }
            }
        }
        return need ? "'" + name.Replace("'", "''") + "'" : name;
    }
}
