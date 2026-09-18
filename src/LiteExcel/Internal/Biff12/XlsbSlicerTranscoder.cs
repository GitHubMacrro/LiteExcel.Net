using System;
using System.Collections.Generic;
using System.Text;

namespace LiteExcel.Internal.Biff12;

/// <summary>
/// xlsb <c>slicerCacheN.bin</c> / <c>slicerN.bin</c> → OOXML（Stage E）。
///
/// slicerCache 记录（实测标定，见 `docs/internal/LiteExcel_StageE_Recon.md`）：
///   <c>0x0433</c> begin / <c>0x0434</c> end；<c>0x0025</c>+<c>0x0C00</c>(16B GUID)+<c>0x0026</c> uid；
///   <c>0x0435</c> name(WS)+sourceName(WS)；<c>0x043D</c> pivotTables(count@0/tabId@4/name@8)；
///   <c>0x043E</c> pivotCacheId(u32)；<c>0x0440</c> levels count；<c>0x0442</c> level(count@0/flag@4/uniqueName@5/sourceCaption)；
///   <c>0x0446</c> items count；<c>0x0448</c> item(flag@0/n@1/c)；<c>0x0449</c> selections count；<c>0x044A</c> selection(flag@0/n@4)。
/// slicer 记录：<c>0x045B</c> begin / <c>0x045C</c> end；<c>0x043B</c> header(13B)+rowHeight(u32)+name/cache/caption(WS)。
/// </summary>
internal static class XlsbSlicerTranscoder
{
    private const int RtBeginSlicerCache = 0x0433;
    private const int RtName = 0x0435;
    private const int RtPivotTables = 0x043D;
    private const int RtPivotCacheId = 0x043E;
    private const int RtLevelsCount = 0x0440;
    private const int RtLevel = 0x0442;
    private const int RtItemsCount = 0x0446;
    private const int RtItem = 0x0448;
    private const int RtSelectionsCount = 0x0449;
    private const int RtSelection = 0x044A;
    private const int RtBeginSlicer = 0x045B;
    private const int RtSlicer = 0x043B;

    private const string SlicerNs = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string McNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private const string Xr10Ns = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10";

    internal sealed class SlicerPivotTable
    {
        public uint TabId;
        public string Name = "";
    }

    internal sealed class SlicerItem
    {
        public string Name = "";
        public string Caption = "";
    }

    internal sealed class SlicerLevel
    {
        public string UniqueName = "";
        public string SourceCaption = "";
        public int Count;
        public readonly List<SlicerItem> Items = new();
    }

    internal sealed class SlicerCacheInfo
    {
        public string Name = "";
        public string SourceName = "";
        public string Uid = "";
        public uint PivotCacheId;
        public readonly List<SlicerPivotTable> PivotTables = new();
        public readonly List<SlicerLevel> Levels = new();
        public readonly List<string> Selections = new();
        public bool SelectionPivotItem;
    }

    internal sealed class SlicerInfo
    {
        public string Name = "";
        public string Cache = "";
        public string Caption = "";
        public string Uid = "";
        public int Level = 1;
        public uint RowHeight;
    }

    public static SlicerCacheInfo ParseCache(byte[] data)
    {
        var info = new SlicerCacheInfo();
        SlicerLevel? level = null;
        foreach (var rec in Biff12Records.ReadAll(data))
        {
            var d = rec.Data;
            switch (rec.Rt)
            {
                case 0x0C00:
                    if (d.Length >= 16) info.Uid = FormatGuid(d, 0);
                    break;
                case RtName:
                {
                    int off = 0;
                    info.Name = ReadWideString(d, ref off);
                    info.SourceName = ReadWideString(d, ref off);
                    break;
                }
                case RtPivotTables:
                    if (d.Length >= 8)
                    {
                        uint count = Biff12Records.ReadU32(d, 0);
                        int off = 4;
                        for (uint i = 0; i < count; i++)
                        {
                            if (off + 4 > d.Length) break;
                            var pt = new SlicerPivotTable { TabId = Biff12Records.ReadU32(d, off) };
                            off += 4;
                            pt.Name = ReadWideString(d, ref off);
                            info.PivotTables.Add(pt);
                        }
                    }
                    break;
                case RtPivotCacheId:
                    if (d.Length >= 4) info.PivotCacheId = Biff12Records.ReadU32(d, 0);
                    break;
                case RtLevel:
                    level = ParseLevel(d);
                    info.Levels.Add(level);
                    break;
                case RtItem:
                    if (level is not null)
                    {
                        int off = 1;
                        var it = new SlicerItem { Name = ReadWideString(d, ref off) };
                        it.Caption = ReadWideString(d, ref off);
                        level.Items.Add(it);
                    }
                    break;
                case RtSelection:
                    if (d.Length >= 4)
                    {
                        info.SelectionPivotItem = (d[0] & 0x01) != 0;
                        int off = 4;
                        var s = ReadWideString(d, ref off);
                        if (s.Length > 0) info.Selections.Add(s);
                    }
                    break;
            }
        }
        return info;
    }

    private static SlicerLevel ParseLevel(byte[] d)
    {
        var lv = new SlicerLevel();
        if (d.Length < 5) return lv;
        lv.Count = (int)Biff12Records.ReadU32(d, 0);
        int off = 5;
        lv.UniqueName = ReadWideString(d, ref off);
        lv.SourceCaption = ReadWideString(d, ref off);
        return lv;
    }

    public static SlicerInfo ParseSlicer(byte[] data)
    {
        var list = ParseSlicers(data);
        return list.Count > 0 ? list[0] : new SlicerInfo();
    }

    /// <summary>一个 slicerN.bin 可含多个 slicer（每个由 0x043B 记录 + 前置 uid 构成）。</summary>
    public static List<SlicerInfo> ParseSlicers(byte[] data)
    {
        var list = new List<SlicerInfo>();
        SlicerInfo? cur = null;
        string pendingUid = "";
        foreach (var rec in Biff12Records.ReadAll(data))
        {
            var d = rec.Data;
            if (rec.Rt == 0x0C00)
            {
                if (d.Length >= 16) pendingUid = FormatGuid(d, 0);
            }
            else if (rec.Rt == RtSlicer && d.Length >= 17)
            {
                cur = new SlicerInfo { Uid = pendingUid, RowHeight = Biff12Records.ReadU32(d, 13) };
                int off = 17;
                cur.Name = ReadWideString(d, ref off);
                cur.Cache = ReadWideString(d, ref off);
                cur.Caption = ReadWideString(d, ref off);
                list.Add(cur);
            }
        }
        return list;
    }

    private static string FormatGuid(byte[] d, int off)
    {
        var g = new Guid(
            Biff12Records.ReadU32(d, off), Biff12Records.ReadU16(d, off + 4), Biff12Records.ReadU16(d, off + 6),
            d[off + 8], d[off + 9], d[off + 10], d[off + 11], d[off + 12], d[off + 13], d[off + 14], d[off + 15]);
        return "{" + g.ToString().ToUpperInvariant() + "}";
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

    public static string ToSlicerXml(IReadOnlyList<SlicerInfo> slicers)
    {
        var sb = new StringBuilder(512);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<slicers xmlns=\"{SlicerNs}\" xmlns:mc=\"{McNs}\" mc:Ignorable=\"x xr10\" " +
                  $"xmlns:x=\"{MainNs}\" xmlns:xr10=\"{Xr10Ns}\">");
        foreach (var s in slicers)
        {
            sb.Append($"<slicer name=\"{Esc(s.Name)}\"");
            if (s.Uid.Length > 0) sb.Append($" xr10:uid=\"{s.Uid}\"");
            sb.Append($" cache=\"{Esc(s.Cache)}\" caption=\"{Esc(s.Caption)}\" level=\"{s.Level}\" rowHeight=\"{s.RowHeight}\"/>");
        }
        sb.Append("</slicers>");
        return sb.ToString();
    }

    public static string ToSlicerCacheXml(SlicerCacheInfo c)
    {
        var sb = new StringBuilder(1024);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<slicerCacheDefinition xmlns=\"{SlicerNs}\" xmlns:mc=\"{McNs}\" mc:Ignorable=\"x xr10\" " +
                  $"xmlns:x=\"{MainNs}\" xmlns:xr10=\"{Xr10Ns}\" name=\"{Esc(c.Name)}\"");
        if (c.Uid.Length > 0) sb.Append($" xr10:uid=\"{c.Uid}\"");
        sb.Append($" sourceName=\"{Esc(c.SourceName)}\">");

        sb.Append("<pivotTables>");
        foreach (var pt in c.PivotTables)
            sb.Append($"<pivotTable tabId=\"{pt.TabId}\" name=\"{Esc(pt.Name)}\"/>");
        sb.Append("</pivotTables>");

        sb.Append("<data><olap");
        if (c.PivotCacheId != 0) sb.Append($" pivotCacheId=\"{c.PivotCacheId}\"");
        sb.Append(">");
        sb.Append($"<levels count=\"{c.Levels.Count}\">");
        foreach (var lv in c.Levels)
        {
            sb.Append($"<level uniqueName=\"{Esc(lv.UniqueName)}\" sourceCaption=\"{Esc(lv.SourceCaption)}\" count=\"{lv.Count}\"");
            if (lv.Items.Count > 0)
            {
                sb.Append("><ranges><range startItem=\"0\">");
                foreach (var it in lv.Items)
                    sb.Append($"<i n=\"{Esc(it.Name)}\" c=\"{Esc(it.Caption)}\"/>");
                sb.Append("</range></ranges></level>");
            }
            else
            {
                sb.Append("/>");
            }
        }
        sb.Append("</levels>");

        if (c.Selections.Count > 0)
        {
            sb.Append($"<selections count=\"{c.Selections.Count}\">");
            foreach (var s in c.Selections)
            {
                if (c.SelectionPivotItem) sb.Append($"<selection pivotItem=\"1\" n=\"{Esc(s)}\"/>");
                else sb.Append($"<selection n=\"{Esc(s)}\"/>");
            }
            sb.Append("</selections>");
        }

        sb.Append("</olap></data></slicerCacheDefinition>");
        return sb.ToString();
    }
}
