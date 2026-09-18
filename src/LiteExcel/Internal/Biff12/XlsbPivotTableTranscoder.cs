using System;
using System.Collections.Generic;
using System.Text;

namespace LiteExcel.Internal.Biff12;

/// <summary>
/// xlsb <c>pivotTableN.bin</c> → <c>pivotTableN.xml</c>（Stage D，进行中）。
///
/// 已标定记录（MS-XLSB 规范 + 真实样本逐字节）：
///   <c>0x0118</c> BrtBeginSXView：头 32 字节（idCache@28），随后 XLWideString name / dataCaption / rowHeaderCaption；
///   <c>0x013A</c> BrtBeginSxLocation：rwFirst/rwLast/colFirst/colLast + firstHeaderRow/firstDataRow/firstDataCol；
///   <c>0x011F</c> BrtBeginSXVDs(csxvds)；<c>0x011D</c> BrtBeginSXVD(20B 固定)；
///   <c>0x011B</c> BrtBeginSXVIs(csxvis)；<c>0x011A</c> BrtBeginSXVI(7B)；
///   <c>0x0135</c> BrtBeginISXVDRws(rowFields)；<c>0x012B</c> BrtBeginSXLIRws + <c>0x0129</c> SXLI(12B) + <c>0x0184</c> ISXVIs(rowItems)；
///   <c>0x013C</c> BrtBeginPivotHierarchies + <c>0x013E</c> BrtPivotHierarchy(6B)；
///   <c>0x0140</c> BrtBeginRowHierarchiesUsage；<c>0x0201</c> BrtTableStyleClient；<c>0x013B</c> BrtEndSXView。
/// 列方向：<c>0x0137</c> BrtBeginISXVDCols(colFields) / <c>0x012D</c> BrtBeginSXLICols + <c>0x0129</c>/<c>0x0184</c>（colItems）；
/// 数据字段：<c>0x0125</c> BrtBeginPivotDataField（isxvdData@0/iiftab@4/df@8/isxvd@12/isxvi@16/ifmt@20/fLoadDisplayName@24/stDisplayName）；
/// <c>0x0142</c> BrtBeginColHierarchiesUsage。其余记录（SXPI/SXDI/formats/extLst）待补。
/// </summary>
internal static class XlsbPivotTableTranscoder
{
    private const int RtBeginPivotTable = 0x0118;
    private const int RtLocation = 0x013A;
    private const int RtPivotFieldsCount = 0x011F;
    private const int RtPivotField = 0x011D;
    private const int RtItemsCount = 0x011B;
    private const int RtItem = 0x011A;
    private const int RtRowFields = 0x0135;
    private const int RtColFields = 0x0137;
    private const int RtRowItemsCount = 0x012B;
    private const int RtRowItemsEnd = 0x012C;
    private const int RtColItemsCount = 0x012D;
    private const int RtColItemsEnd = 0x012E;
    private const int RtLine = 0x0129;
    private const int RtLineEntries = 0x0184;
    private const int RtDataFieldsCount = 0x0127;
    private const int RtDataField = 0x0125;
    private const int RtHierarchiesCount = 0x013C;
    private const int RtHierarchy = 0x013E;
    private const int RtRowHierarchyUsage = 0x0140;
    private const int RtColHierarchyUsage = 0x0142;
    private const int RtTableStyleClient = 0x0201;

    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    internal sealed class PivotItem
    {
        public int ItemType;
        public ushort Flags;
        public int CacheIndex;
        public string? DisplayName;
    }

    internal sealed class PivotField
    {
        public int SxAxis;
        public ushort SubtotalFlags;
        public byte FlagsNu;
        public uint FlagsVt;
        public string? Name;
        public string? SubtotalCaption;
        public readonly List<PivotItem> Items = new();
    }

    internal sealed class PivotLine
    {
        public int RowIndex;
        public byte ItemType;
        public int DataIndex;
        public readonly List<uint> Entries = new();
    }

    internal sealed class DataField
    {
        public string Name = "";
        public int Field;
        public int BaseField;
        public int BaseItem;
    }

    internal sealed class PivotTableInfo
    {
        public uint CacheId;
        public string Name = "";
        public string? DataCaption;
        public string? Tag;
        public string? RowHeaderCaption;
        public string? PivotStyle;
        public int PivotFieldCount;
        public (int RwFirst, int RwLast, int ColFirst, int ColLast) Location;
        public bool HasLocation;
        public int FirstHeaderRow;
        public int FirstDataRow;
        public int FirstDataCol;
        public readonly List<PivotField> Fields = new();
        public readonly List<int> RowFields = new();
        public readonly List<PivotLine> RowItems = new();
        public readonly List<int> ColFields = new();
        public readonly List<PivotLine> ColItems = new();
        public readonly List<DataField> DataFields = new();
        public readonly List<ushort> Hierarchies = new();
        public readonly List<int> RowHierarchyUsage = new();
        public readonly List<int> ColHierarchyUsage = new();
    }

    public static PivotTableInfo Parse(byte[] data)
    {
        var info = new PivotTableInfo();
        PivotField? curField = null;
        PivotLine? curLine = null;
        List<PivotLine>? curLines = null;
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
                case RtPivotField:
                    curField = ParseField(d);
                    info.Fields.Add(curField);
                    break;
                case RtItem:
                    if (curField is not null) curField.Items.Add(ParseItem(d));
                    break;
                case RtRowFields:
                    ParseIndexArray(d, info.RowFields);
                    break;
                case RtColFields:
                    ParseIndexArray(d, info.ColFields);
                    break;
                case RtRowItemsCount:
                    curLines = info.RowItems;
                    curLine = null;
                    break;
                case RtColItemsCount:
                    curLines = info.ColItems;
                    curLine = null;
                    break;
                case RtRowItemsEnd:
                case RtColItemsEnd:
                    curLines = null;
                    curLine = null;
                    break;
                case RtLine:
                    curLine = ParseLine(d);
                    curLines?.Add(curLine);
                    break;
                case RtLineEntries:
                    if (curLine is not null)
                        for (int o = 0; o + 4 <= d.Length; o += 4)
                            curLine.Entries.Add(Biff12Records.ReadU32(d, o));
                    break;
                case RtDataField:
                    info.DataFields.Add(ParseDataField(d));
                    break;
                case RtHierarchy:
                    if (d.Length >= 2) info.Hierarchies.Add(Biff12Records.ReadU16(d, 0));
                    break;
                case RtRowHierarchyUsage:
                    ParseIndexArray(d, info.RowHierarchyUsage);
                    break;
                case RtColHierarchyUsage:
                    ParseIndexArray(d, info.ColHierarchyUsage);
                    break;
                case RtTableStyleClient:
                {
                    int off = 2;
                    var style = ReadWideString(d, ref off);
                    if (style.Length > 0) info.PivotStyle = style;
                    break;
                }
            }
        }
        return info;
    }

    private static PivotField ParseField(byte[] d)
    {
        var f = new PivotField();
        if (d.Length < 20) return f;
        f.SxAxis = d[0];
        f.SubtotalFlags = Biff12Records.ReadU16(d, 1);
        f.FlagsNu = d[3];
        f.FlagsVt = Biff12Records.ReadU32(d, 8);
        int off = 20;
        if ((f.FlagsNu & 0x20) != 0) f.Name = ReadWideString(d, ref off);
        if ((f.FlagsNu & 0x40) != 0) f.SubtotalCaption = ReadWideString(d, ref off);
        return f;
    }

    private static PivotItem ParseItem(byte[] d)
    {
        var it = new PivotItem();
        if (d.Length < 7) return it;
        it.ItemType = d[0];
        it.Flags = Biff12Records.ReadU16(d, 1);
        it.CacheIndex = Biff12Records.ReadS32(d, 3);
        int off = 7;
        if ((it.Flags & 0x10) != 0) it.DisplayName = ReadWideString(d, ref off);
        return it;
    }

    private static PivotLine ParseLine(byte[] d)
    {
        var line = new PivotLine();
        if (d.Length < 12) return line;
        line.RowIndex = Biff12Records.ReadU16(d, 0);
        line.ItemType = d[2];
        line.DataIndex = Biff12Records.ReadS32(d, 8);
        return line;
    }

    private static void ParseIndexArray(byte[] d, List<int> target)
    {
        if (d.Length < 4) return;
        int count = (int)Biff12Records.ReadU32(d, 0);
        for (int i = 0; i < count; i++)
        {
            int off = 4 + i * 4;
            if (off + 4 > d.Length) break;
            target.Add(Biff12Records.ReadS32(d, off));
        }
    }

    private static DataField ParseDataField(byte[] d)
    {
        var df = new DataField();
        if (d.Length < 25) return df;
        df.Field = Biff12Records.ReadS32(d, 0);
        df.BaseField = Biff12Records.ReadS32(d, 12);
        df.BaseItem = Biff12Records.ReadS32(d, 16);
        int off = 25;
        if (d[24] != 0) df.Name = ReadWideString(d, ref off);
        return df;
    }

    private static void ParseBegin(byte[] d, PivotTableInfo info)
    {
        if (d.Length < 32) return;
        info.CacheId = Biff12Records.ReadU32(d, 28);
        var strs = ReadTailStrings(d, 32);
        if (strs.Count == 0) return;
        info.Name = strs[0];
        int i = 1;
        if (i < strs.Count && !IsGuid(strs[i])) { info.DataCaption = strs[i]; i++; }
        if (i < strs.Count && IsGuid(strs[i])) { info.Tag = strs[i]; i++; }
        if (i < strs.Count) info.RowHeaderCaption = strs[i];
    }

    private static List<string> ReadTailStrings(byte[] d, int off)
    {
        var list = new List<string>();
        while (off + 4 <= d.Length)
        {
            uint cch = Biff12Records.ReadU32(d, off);
            if (cch == 0 || cch > (uint)((d.Length - off - 4) / 2)) break;
            var s = Encoding.Unicode.GetString(d, off + 4, (int)cch * 2);
            list.Add(s);
            off += 4 + (int)cch * 2;
        }
        return list;
    }

    private static bool IsGuid(string s)
    {
        if (s.Length != 36) return false;
        return s[8] == '-' && s[13] == '-' && s[18] == '-' && s[23] == '-';
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

    private static bool Bit(uint v, int i) => (v & (1u << i)) != 0;

    private static string ItemTypeName(int t) => t switch
    {
        0x01 => "default",
        0x02 => "sum",
        0x03 => "countA",
        0x04 => "avg",
        0x05 => "max",
        0x06 => "min",
        0x07 => "product",
        0x08 => "count",
        0x09 => "stdDev",
        0x0A => "stdDevP",
        0x0B => "var",
        0x0C => "varP",
        0x0D => "grand",
        0x0E => "blank",
        _ => "",
    };

    /// <summary>生成 pivotTable 全文（核心段；colFields/colItems/SXPI/SXDI 等待补）。</summary>
    public static string ToXml(PivotTableInfo p)
    {
        var sb = new StringBuilder(1024);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<pivotTableDefinition xmlns=\"{MainNs}\" name=\"{Esc(p.Name)}\" cacheId=\"{p.CacheId}\" " +
                  "applyNumberFormats=\"0\" applyBorderFormats=\"0\" applyFontFormats=\"0\" " +
                  "applyPatternFormats=\"0\" applyAlignmentFormats=\"0\" applyWidthHeightFormats=\"1\"");
        if (p.DataCaption is not null) sb.Append($" dataCaption=\"{Esc(p.DataCaption)}\"");
        if (p.Tag is not null) sb.Append($" tag=\"{Esc(p.Tag)}\"");
        if (p.RowHeaderCaption is not null) sb.Append($" rowHeaderCaption=\"{Esc(p.RowHeaderCaption)}\"");
        sb.Append('>');

        if (p.HasLocation)
        {
            var (rf, rl, cf, cl) = p.Location;
            sb.Append($"<location ref=\"{A1(rf, cf)}:{A1(rl, cl)}\" firstHeaderRow=\"{p.FirstHeaderRow}\" " +
                      $"firstDataRow=\"{p.FirstDataRow}\" firstDataCol=\"{p.FirstDataCol}\"/>");
        }

        AppendFields(sb, p);
        AppendFields(sb, "rowFields", p.RowFields);
        AppendLines(sb, "rowItems", p.RowItems);
        AppendFields(sb, "colFields", p.ColFields);
        AppendLines(sb, "colItems", p.ColItems);
        AppendDataFields(sb, p);

        if (p.Hierarchies.Count > 0)
        {
            sb.Append($"<pivotHierarchies count=\"{p.Hierarchies.Count}\">");
            foreach (var h in p.Hierarchies) AppendHierarchy(sb, h);
            sb.Append("</pivotHierarchies>");
        }

        if (p.PivotStyle is not null)
            sb.Append($"<pivotTableStyleInfo name=\"{Esc(p.PivotStyle)}\" showRowHeaders=\"1\" " +
                      "showColHeaders=\"1\" showRowStripes=\"0\" showColStripes=\"0\" showLastColumn=\"1\"/>");

        if (p.RowHierarchyUsage.Count > 0)
        {
            sb.Append($"<rowHierarchiesUsage count=\"{p.RowHierarchyUsage.Count}\">");
            foreach (var u in p.RowHierarchyUsage)
                sb.Append($"<rowHierarchyUsage hierarchyUsage=\"{u}\"/>");
            sb.Append("</rowHierarchiesUsage>");
        }

        if (p.ColHierarchyUsage.Count > 0)
        {
            sb.Append($"<colHierarchiesUsage count=\"{p.ColHierarchyUsage.Count}\">");
            foreach (var u in p.ColHierarchyUsage)
                sb.Append($"<colHierarchyUsage hierarchyUsage=\"{u}\"/>");
            sb.Append("</colHierarchiesUsage>");
        }

        sb.Append("</pivotTableDefinition>");
        return sb.ToString();
    }

    private static void AppendDataFields(StringBuilder sb, PivotTableInfo p)
    {
        if (p.DataFields.Count == 0) return;
        sb.Append($"<dataFields count=\"{p.DataFields.Count}\">");
        foreach (var df in p.DataFields)
            sb.Append($"<dataField name=\"{Esc(df.Name)}\" fld=\"{df.Field}\" baseField=\"{df.BaseField}\" baseItem=\"{df.BaseItem}\"/>");
        sb.Append("</dataFields>");
    }

    private static void AppendFields(StringBuilder sb, PivotTableInfo p)
    {
        if (p.Fields.Count == 0) return;
        sb.Append($"<pivotFields count=\"{p.Fields.Count}\">");
        foreach (var f in p.Fields) AppendField(sb, f);
        sb.Append("</pivotFields>");
    }

    private static void AppendField(StringBuilder sb, PivotField f)
    {
        sb.Append("<pivotField");
        if (f.Name is not null) sb.Append($" name=\"{Esc(f.Name)}\"");
        if ((f.SxAxis & 0x08) != 0) sb.Append(" dataField=\"1\"");
        else if ((f.SxAxis & 0x01) != 0) sb.Append(" axis=\"axisRow\"");
        else if ((f.SxAxis & 0x02) != 0) sb.Append(" axis=\"axisCol\"");
        else if ((f.SxAxis & 0x04) != 0) sb.Append(" axis=\"axisPage\"");
        if ((f.FlagsNu & 0x10) == 0) sb.Append(" compact=\"0\"");
        if ((f.FlagsNu & 0x01) != 0) sb.Append(" allDrilled=\"1\"");
        if (!Bit(f.FlagsVt, 5)) sb.Append(" showAll=\"0\"");
        if ((f.FlagsNu & 0x80) != 0) sb.Append(" dataSourceSort=\"1\"");
        if (!Bit(f.FlagsVt, 6)) sb.Append(" outline=\"0\"");
        if (!Bit(f.FlagsVt, 8)) sb.Append(" subtotalTop=\"0\"");
        if ((f.SubtotalFlags & 0x01) == 0) sb.Append(" defaultSubtotal=\"0\"");
        if (Bit(f.FlagsVt, 7)) sb.Append(" insertBlankRow=\"1\"");
        if (!Bit(f.FlagsVt, 0)) sb.Append(" dragToRow=\"0\"");
        if (!Bit(f.FlagsVt, 1)) sb.Append(" dragToCol=\"0\"");
        if (!Bit(f.FlagsVt, 2)) sb.Append(" dragToPage=\"0\"");
        if (!Bit(f.FlagsVt, 3)) sb.Append(" dragToHide=\"0\"");
        if (!Bit(f.FlagsVt, 4)) sb.Append(" dragToData=\"0\"");
        AppendSubtotal(sb, "sumSubtotal", f.SubtotalFlags, 1);
        AppendSubtotal(sb, "countASubtotal", f.SubtotalFlags, 2);
        AppendSubtotal(sb, "avgSubtotal", f.SubtotalFlags, 3);
        AppendSubtotal(sb, "maxSubtotal", f.SubtotalFlags, 4);
        AppendSubtotal(sb, "minSubtotal", f.SubtotalFlags, 5);
        AppendSubtotal(sb, "productSubtotal", f.SubtotalFlags, 6);
        AppendSubtotal(sb, "countSubtotal", f.SubtotalFlags, 7);
        AppendSubtotal(sb, "stdDevSubtotal", f.SubtotalFlags, 8);
        AppendSubtotal(sb, "stdDevPSubtotal", f.SubtotalFlags, 9);
        AppendSubtotal(sb, "varSubtotal", f.SubtotalFlags, 10);
        AppendSubtotal(sb, "varPSubtotal", f.SubtotalFlags, 11);
        if (Bit(f.FlagsVt, 24)) sb.Append(" defaultAttributeDrillState=\"1\"");

        if (f.Items.Count > 0)
        {
            sb.Append($"><items count=\"{f.Items.Count}\">");
            foreach (var it in f.Items) AppendItem(sb, it);
            sb.Append("</items></pivotField>");
        }
        else
        {
            sb.Append("/>");
        }
    }

    private static void AppendSubtotal(StringBuilder sb, string name, ushort flags, int bit)
    {
        if ((flags & (1 << bit)) != 0) sb.Append($" {name}=\"1\"");
    }

    private static void AppendItem(StringBuilder sb, PivotItem it)
    {
        var t = ItemTypeName(it.ItemType);
        if (t.Length > 0) sb.Append($"<item t=\"{t}\"");
        else sb.Append($"<item x=\"{it.CacheIndex}\"");
        if ((it.Flags & 0x01) != 0) sb.Append(" h=\"1\"");
        if ((it.Flags & 0x02) != 0) sb.Append(" sd=\"1\"");
        if ((it.Flags & 0x04) != 0) sb.Append(" f=\"1\"");
        if ((it.Flags & 0x08) != 0) sb.Append(" m=\"1\"");
        if ((it.Flags & 0x10) != 0) sb.Append(" d=\"1\"");
        if ((it.Flags & 0x20) != 0) sb.Append(" r=\"1\"");
        if (it.DisplayName is not null) sb.Append($" v=\"{Esc(it.DisplayName)}\"");
        sb.Append("/>");
    }

    private static void AppendFields(StringBuilder sb, string tag, List<int> fields)
    {
        if (fields.Count == 0) return;
        var emit = new List<int>();
        foreach (var x in fields)
            if (x != -1) emit.Add(x);
        if (emit.Count == 0) return;
        sb.Append($"<{tag} count=\"{emit.Count}\">");
        foreach (var x in emit) sb.Append($"<field x=\"{x}\"/>");
        sb.Append($"</{tag}>");
    }

    private static void AppendLines(StringBuilder sb, string tag, List<PivotLine> lines)
    {
        if (lines.Count == 0) return;
        string indexAttr = tag == "rowItems" ? "r" : "i";
        sb.Append($"<{tag} count=\"{lines.Count}\">");
        foreach (var line in lines)
        {
            if (line.RowIndex != 0) sb.Append($"<i {indexAttr}=\"{line.RowIndex}\">");
            else sb.Append("<i>");
            foreach (var e in line.Entries)
                sb.Append(e == 0 ? "<x/>" : $"<x v=\"{e}\"/>");
            sb.Append("</i>");
        }
        sb.Append($"</{tag}>");
    }

    private static void AppendHierarchy(StringBuilder sb, ushort flags)
    {
        sb.Append("<pivotHierarchy");
        if ((flags & 0x10) == 0) sb.Append(" dragToRow=\"0\"");
        if ((flags & 0x20) == 0) sb.Append(" dragToCol=\"0\"");
        if ((flags & 0x40) == 0) sb.Append(" dragToPage=\"0\"");
        if ((flags & 0x80) != 0) sb.Append(" dragToData=\"1\"");
        sb.Append("/>");
    }
}
