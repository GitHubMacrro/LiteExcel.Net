using System;
using System.Collections.Generic;
using System.Text;

namespace LiteExcel.Internal.Biff12;

/// <summary>
/// xlsb <c>pivotTableN.bin</c> → <c>pivotTableN.xml</c>（Stage D，进行中）。
///
/// 已标定记录（真实样本）：
///   <c>0x0118</c> BrtBeginPivotTable：头 32 字节（cacheId@28），随后 XLWideString name / dataCaption / rowHeaderCaption；
///   <c>0x013A</c> BrtBeginPivotTableLocation：rwFirst(u32)/rwLast/colFirst/colLast + firstHeaderRow/firstDataRow/firstDataCol；
///   <c>0x011F</c> pivotFields 计数；<c>0x0119</c>/<c>0x011A</c> 字段；
///   <c>0x0129</c>/<c>0x012A</c> 行项；<c>0x0201</c> BrtTableStyleClient（数据透视表样式）。
/// 其余记录（62 种类型）待补。
/// </summary>
internal static class XlsbPivotTableTranscoder
{
    private const int RtBeginPivotTable = 0x0118;
    private const int RtLocation = 0x013A;
    private const int RtPivotFieldsCount = 0x011F;
    private const int RtTableStyleClient = 0x0201;

    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    internal sealed class PivotTableInfo
    {
        public uint CacheId;
        public string Name = "";
        public string? DataCaption;
        public string? RowHeaderCaption;
        public string? PivotStyle;
        public int PivotFieldCount;
        public (int RwFirst, int RwLast, int ColFirst, int ColLast) Location;
        public bool HasLocation;
        public int FirstHeaderRow;
        public int FirstDataRow;
        public int FirstDataCol;
    }

    public static PivotTableInfo Parse(byte[] data)
    {
        var info = new PivotTableInfo();
        foreach (var rec in Biff12Records.ReadAll(data))
        {
            var d = rec.Data;
            switch (rec.Rt)
            {
                case RtBeginPivotTable:
                    ParseBegin(d, info);
                    break;
                case RtLocation:
                    if (d.Length >= 28)
                    {
                        int rwFirst = (int)Biff12Records.ReadU32(d, 0);
                        int colFirst = (int)Biff12Records.ReadU32(d, 8);
                        info.Location = (rwFirst, (int)Biff12Records.ReadU32(d, 4),
                                         colFirst, (int)Biff12Records.ReadU32(d, 12));
                        // [4]/[5] 为 rwFirst + firstHeaderRow/firstDataRow；[6] 为 colFirst + firstDataCol（真实样本标定）。
                        info.FirstHeaderRow = (int)Biff12Records.ReadU32(d, 16) - rwFirst;
                        info.FirstDataRow = (int)Biff12Records.ReadU32(d, 20) - rwFirst;
                        info.FirstDataCol = (int)Biff12Records.ReadU32(d, 24) - colFirst;
                        info.HasLocation = true;
                    }
                    break;
                case RtPivotFieldsCount:
                    if (d.Length >= 4) info.PivotFieldCount = (int)Biff12Records.ReadU32(d, 0);
                    break;
                case RtTableStyleClient:
                {
                    // 32(2) + XLWideString name
                    int off = 2;
                    var style = ReadWideString(d, ref off);
                    if (style.Length > 0) info.PivotStyle = style;
                    break;
                }
            }
        }
        return info;
    }

    private static void ParseBegin(byte[] d, PivotTableInfo info)
    {
        if (d.Length < 32) return;
        info.CacheId = Biff12Records.ReadU32(d, 28);
        int off = 32;
        info.Name = ReadWideString(d, ref off);
        info.DataCaption = ReadWideString(d, ref off);
        info.RowHeaderCaption = ReadWideString(d, ref off);
    }

    private static string ReadWideString(byte[] d, ref int off)
    {
        if (off + 4 > d.Length) return "";
        uint cch = Biff12Records.ReadU32(d, off);
        off += 4;
        if (cch == 0 || cch > (uint)((d.Length - off) / 2)) return "";
        var s = Encoding.Unicode.GetString(d, off, (int)cch * 2);
        off += (int)cch * 2;
        return s;
    }

    private static string Esc(string s) => s
        .Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");

    private static string A1(int row, int col)
    {
        var name = RgceDecoder.ColumnName(col);
        return name + (row + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    /// <summary>生成 pivotTable 全文（核心段；字段/项结构待补）。</summary>
    public static string ToXml(PivotTableInfo p)
    {
        var sb = new StringBuilder(512);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<pivotTableDefinition xmlns=\"{MainNs}\" name=\"{Esc(p.Name)}\" cacheId=\"{p.CacheId}\" " +
                  "applyNumberFormats=\"0\" applyBorderFormats=\"0\" applyFontFormats=\"0\" " +
                  "applyPatternFormats=\"0\" applyAlignmentFormats=\"0\" applyWidthHeightFormats=\"1\"");
        if (p.DataCaption is not null) sb.Append($" dataCaption=\"{Esc(p.DataCaption)}\"");
        if (p.RowHeaderCaption is not null) sb.Append($" rowHeaderCaption=\"{Esc(p.RowHeaderCaption)}\"");
        sb.Append('>');
        if (p.HasLocation)
        {
            var (rf, rl, cf, cl) = p.Location;
            sb.Append($"<location ref=\"{A1(rf, cf)}:{A1(rl, cl)}\" firstHeaderRow=\"{p.FirstHeaderRow}\" " +
                      $"firstDataRow=\"{p.FirstDataRow}\" firstDataCol=\"{p.FirstDataCol}\"/>");
        }
        sb.Append("</pivotTableDefinition>");
        return sb.ToString();
    }
}
