using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace LiteExcel.Internal.Biff12;

/// <summary>
/// xlsb <c>pivotCacheDefinitionN.bin</c> → <c>pivotCacheDefinitionN.xml</c>（Stage D，进行中）。
///
/// 记录层级（真实样本逐字节标定，见 `docs/internal/LiteExcel_StageD_Recon.md`）：
///   <c>0x00B3</c> 头（版本 / refreshedDate(f64) / refreshedBy）→ <c>0x00B9</c> cacheSource（type + connectionId）
///   → <c>0x00B5</c> cacheFields 计数 → [<c>0x00B7</c> cacheField（hierarchy/level/numFmtId/name/caption）
///   + <c>0x00BD</c> sharedItems 头 + <c>0x0018</c> 值 + <c>0x00BE</c> + <c>0x00B8</c>] → <c>0x082C</c> cacheHierarchies …
///
/// 说明：本类当前实现「头 / cacheSource / cacheFields / sharedItems」核心段；
/// cacheHierarchies（0x082C / 0x00C5 / 0x00C7 / 0x083C / 0x010D-0x0114 / 0x01E6-0x01ED）为数据模型 OLAP 结构，后续补齐。
/// </summary>
internal static class XlsbPivotCacheTranscoder
{
    private const int RtCacheDef = 0x00B3;
    private const int RtCacheSource = 0x00B9;
    private const int RtBeginCacheFields = 0x00B5;
    private const int RtCacheField = 0x00B7;
    private const int RtSharedItemsHead = 0x00BD;
    private const int RtSharedItemStr = 0x0018;

    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    internal sealed class CacheFieldInfo
    {
        public string Name = "";
        public string? Caption;
        public uint NumFmtId;
        public int Hierarchy;
        public int Level;
        public readonly List<string> SharedStrings = new();
    }

    internal sealed class PivotCacheInfo
    {
        public int RefreshedVersion;
        public int MinRefreshableVersion;
        public int CreatedVersion;
        public string? RefreshedBy;
        public double? RefreshedDate;
        public int SourceType;       // 1 = external（数据模型/OLAP）
        public uint ConnectionId;
        public readonly List<CacheFieldInfo> Fields = new();
    }

    public static PivotCacheInfo Parse(byte[] data)
    {
        var info = new PivotCacheInfo();
        CacheFieldInfo? field = null;
        foreach (var rec in Biff12Records.ReadAll(data))
        {
            var d = rec.Data;
            switch (rec.Rt)
            {
                case RtCacheDef:
                    ParseHeader(d, info);
                    break;
                case RtCacheSource:
                    if (d.Length >= 8) { info.SourceType = (int)Biff12Records.ReadU32(d, 0); info.ConnectionId = Biff12Records.ReadU32(d, 4); }
                    break;
                case RtCacheField:
                    field = ParseField(d);
                    info.Fields.Add(field);
                    break;
                case RtSharedItemStr:
                    if (field is not null)
                    {
                        int off = 0;
                        var s = ReadWideString(d, ref off);
                        if (s.Length > 0) field.SharedStrings.Add(s);
                    }
                    break;
            }
        }
        return info;
    }

    private static void ParseHeader(byte[] d, PivotCacheInfo info)
    {
        if (d.Length < 25) return;
        info.RefreshedVersion = d[0];
        info.MinRefreshableVersion = d[1];
        info.CreatedVersion = d[2];
        // b3 为标志；b4-7 保留；b8-15 = refreshedDate(f64 LE)
        info.RefreshedDate = BitConverter.ToDouble(d, 8);
        // b16-20 保留/标志；随后 refreshedBy(XLWideString)
        int off = 21;
        var by = ReadWideString(d, ref off);
        if (by.Length > 0) info.RefreshedBy = by;
    }

    private static CacheFieldInfo ParseField(byte[] d)
    {
        var f = new CacheFieldInfo();
        if (d.Length < 24) return f;
        f.Hierarchy = (int)Biff12Records.ReadU32(d, 8);
        f.Level = (int)Biff12Records.ReadU32(d, 12);
        f.NumFmtId = Biff12Records.ReadU32(d, 16);
        int off = 20;
        f.Name = ReadWideString(d, ref off);
        if (off < d.Length)
        {
            var caption = ReadWideString(d, ref off);
            if (caption.Length > 0) f.Caption = caption;
        }
        return f;
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

    /// <summary>生成 pivotCacheDefinition 全文（核心段；hierarchies 待补）。</summary>
    public static string ToXml(PivotCacheInfo c)
    {
        var sb = new StringBuilder(1024);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<pivotCacheDefinition xmlns=\"{MainNs}\" r:id=\"rId1\" " +
                  $"refreshedVersion=\"{c.RefreshedVersion}\" minRefreshableVersion=\"{c.MinRefreshableVersion}\" " +
                  $"createdVersion=\"{c.CreatedVersion}\" recordCount=\"0\"");
        if (c.RefreshedBy is not null) sb.Append($" refreshedBy=\"{Esc(c.RefreshedBy)}\"");
        if (c.RefreshedDate is { } dt && dt > 0)
            sb.Append($" refreshedDate=\"{dt.ToString("0.###############", CultureInfo.InvariantCulture)}\"");
        sb.Append('>');

        if (c.SourceType == 1)
            sb.Append($"<cacheSource type=\"external\"><connection r:id=\"rId1\"/></cacheSource>");
        else
            sb.Append("<cacheSource type=\"worksheet\"/>");

        sb.Append($"<cacheFields count=\"{c.Fields.Count}\">");
        foreach (var f in c.Fields)
        {
            sb.Append($"<cacheField name=\"{Esc(f.Name)}\"");
            if (f.Caption is not null) sb.Append($" caption=\"{Esc(f.Caption)}\"");
            sb.Append($" numFmtId=\"{f.NumFmtId}\" hierarchy=\"{f.Hierarchy}\" level=\"{f.Level}\"");
            if (f.SharedStrings.Count > 0)
            {
                sb.Append($"><sharedItems count=\"{f.SharedStrings.Count}\">");
                foreach (var v in f.SharedStrings)
                    sb.Append($"<s v=\"{Esc(v)}\"/>");
                sb.Append("</sharedItems></cacheField>");
            }
            else
            {
                sb.Append("/>");
            }
        }
        sb.Append("</cacheFields>");
        sb.Append("</pivotCacheDefinition>");
        return sb.ToString();
    }
}
