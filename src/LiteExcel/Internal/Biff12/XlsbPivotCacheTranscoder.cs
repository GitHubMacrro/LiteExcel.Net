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
    private const int RtSharedItemDate = 0x0019;
    private const int RtCacheHierarchy = 0x00C5;
    private const int RtEndCacheHierarchy = 0x00C6;
    private const int RtFieldsUsage = 0x00C7;
    private const int RtEndFieldsUsage = 0x00C8;
    private const int RtDimensionsCount = 0x0111;
    private const int RtDimension = 0x0113;
    private const int RtMeasureGroupsCount = 0x01E6;
    private const int RtMeasureGroup = 0x01EA;
    private const int RtMapsCount = 0x01E8;
    private const int RtMap = 0x01EC;

    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    internal sealed class CacheFieldInfo
    {
        public string Name = "";
        public string? Caption;
        public uint NumFmtId;
        public int Hierarchy;
        public int Level;
        public bool HasSharedItemsHead;
        public byte SharedItemsFlags;
        public double SharedMin;
        public double SharedMax;
        public readonly List<string> SharedStrings = new();
        public readonly List<string> SharedDates = new();
    }

    internal sealed class CacheHierarchyInfo
    {
        public string UniqueName = "";
        public string Caption = "";
        public bool Measure;
        public bool OneField;
        public bool Attribute;
        public int Count;
        public int MemberValueDatatype;
        public string? DimensionUniqueName;
        public string? DefaultMemberUniqueName;
        public string? AllUniqueName;
        public string? DisplayFolder;
        public string? MeasureGroup;
        public readonly List<int> FieldsUsage = new();
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
        public readonly List<CacheHierarchyInfo> Hierarchies = new();
        public readonly List<DimensionInfo> Dimensions = new();
        public readonly List<MeasureGroupInfo> MeasureGroups = new();
        public readonly List<(uint MeasureGroup, uint Dimension)> Maps = new();
    }

    internal sealed class DimensionInfo
    {
        public bool Measure;
        public string Name = "";
        public string UniqueName = "";
        public string Caption = "";
    }

    internal sealed class MeasureGroupInfo
    {
        public string Name = "";
        public string Caption = "";
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
                case RtSharedItemsHead:
                    if (field is not null)
                    {
                        field.HasSharedItemsHead = true;
                        if (d.Length >= 1) field.SharedItemsFlags = d[0];
                        if (d.Length >= 22) { field.SharedMin = BitConverter.ToDouble(d, 6); field.SharedMax = BitConverter.ToDouble(d, 14); }
                    }
                    break;
                case RtSharedItemStr:
                    if (field is not null)
                    {
                        int off = 0;
                        var s = ReadWideString(d, ref off);
                        if (s.Length > 0) field.SharedStrings.Add(s);
                    }
                    break;
                case RtSharedItemDate:
                    if (field is not null && d.Length >= 8)
                    {
                        int year = Biff12Records.ReadU16(d, 0);
                        int mon = d[2], day = d[4], hour = d[5], min = d[6], sec = d[7];
                        field.SharedDates.Add($"{year:D4}-{mon:D2}-{day:D2}T{hour:D2}:{min:D2}:{sec:D2}");
                    }
                    break;
                case RtCacheHierarchy:
                {
                    var h = ParseHierarchy(d);
                    if (h is not null) info.Hierarchies.Add(h);
                    break;
                }
                case RtFieldsUsage:
                    if (info.Hierarchies.Count > 0) ParseFieldsUsage(d, info.Hierarchies[info.Hierarchies.Count - 1]);
                    break;
                case RtDimension:
                    info.Dimensions.Add(ParseDimension(d));
                    break;
                case RtMeasureGroup:
                    info.MeasureGroups.Add(ParseMeasureGroup(d));
                    break;
                case RtMap:
                    if (d.Length >= 8) info.Maps.Add((Biff12Records.ReadU32(d, 0), Biff12Records.ReadU32(d, 4)));
                    break;
            }
        }
        return info;
    }

    /// <summary>BrtBeginPivotCacheHierarchy(0x00C5)：flags(1) measureFlags(1) count(1) … memberValueDatatype(1)@15
    /// + strings@17。字符串布局按类型区分：
    /// 属性层级（flags0 bit2=0x04 或 f1!=0）：uniqueName / caption / dimensionUniqueName / defaultMemberUniqueName / allUniqueName / displayFolder；
    /// 度量层级（flags0 bit0=0x01 且 f1 bit4=0x10）：uniqueName / caption / displayFolder / measureGroup。</summary>
    private static CacheHierarchyInfo? ParseHierarchy(byte[] d)
    {
        if (d.Length < 21) return null;
        bool isMeasure = (d[0] & 0x01) != 0 && (d[1] & 0x10) != 0;
        var h = new CacheHierarchyInfo
        {
            Measure = isMeasure,
            OneField = (d[0] & 0x10) != 0,
            Attribute = !isMeasure && ((d[0] & 0x04) != 0 || d[1] != 0),
            Count = d[2],
            MemberValueDatatype = d[15],
        };
        int off = 17;
        h.UniqueName = ReadWideString(d, ref off);
        h.Caption = ReadWideString(d, ref off);
        if (isMeasure)
        {
            h.DisplayFolder = NullIfEmpty(ReadWideString(d, ref off));
            h.MeasureGroup = NullIfEmpty(ReadWideString(d, ref off));
        }
        else
        {
            h.DimensionUniqueName = NullIfEmpty(ReadWideString(d, ref off));
            h.DefaultMemberUniqueName = NullIfEmpty(ReadWideString(d, ref off));
            h.AllUniqueName = NullIfEmpty(ReadWideString(d, ref off));
            h.DisplayFolder = NullIfEmpty(ReadWideString(d, ref off));
        }
        return h;
    }

    /// <summary>BrtBeginPivotCacheHierarchyFieldsUsage(0x00C7)：count(u32) + count×fieldUsage(i32)。</summary>
    private static void ParseFieldsUsage(byte[] d, CacheHierarchyInfo h)
    {
        if (d.Length < 4) return;
        int count = (int)Biff12Records.ReadU32(d, 0);
        for (int i = 0; i < count; i++)
        {
            int off = 4 + i * 4;
            if (off + 4 > d.Length) break;
            h.FieldsUsage.Add(Biff12Records.ReadS32(d, off));
        }
    }

    private static string? NullIfEmpty(string s) => s.Length == 0 ? null : s;

    /// <summary>BrtBeginPCDSDimension(0x0113)：flags(u8) + name + uniqueName + caption（XLWideString）。flags bit0 = measure。</summary>
    private static DimensionInfo ParseDimension(byte[] d)
    {
        var dim = new DimensionInfo();
        if (d.Length < 1) return dim;
        dim.Measure = (d[0] & 0x01) != 0;
        int off = 1;
        dim.Name = ReadWideString(d, ref off);
        dim.UniqueName = ReadWideString(d, ref off);
        dim.Caption = ReadWideString(d, ref off);
        return dim;
    }

    /// <summary>BrtBeginPCDSMeasureGroup(0x01EA)：flags(u8) + name + caption（XLWideString）。</summary>
    private static MeasureGroupInfo ParseMeasureGroup(byte[] d)
    {
        var mg = new MeasureGroupInfo();
        if (d.Length < 1) return mg;
        int off = 1;
        mg.Name = ReadWideString(d, ref off);
        mg.Caption = ReadWideString(d, ref off);
        return mg;
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

    private static string FmtOle(double ole)
        => DateTime.FromOADate(ole).ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);

    /// <summary>数据字段度量（[Measures].[Sum of X]）对应的属性层级索引；无法匹配返回 -1。</summary>
    private static int AggregatedColumn(PivotCacheInfo c, CacheHierarchyInfo measure)
    {
        const string prefix = "[Measures].[Sum of ";
        if (!measure.UniqueName.StartsWith(prefix, StringComparison.Ordinal)) return -1;
        int end = measure.UniqueName.IndexOf(']', prefix.Length);
        if (end < 0) return -1;
        var field = measure.UniqueName.Substring(prefix.Length, end - prefix.Length);
        var dim = measure.MeasureGroup is null ? null : "[" + measure.MeasureGroup + "]";
        for (int i = 0; i < c.Hierarchies.Count; i++)
        {
            var h = c.Hierarchies[i];
            if (h.Measure) continue;
            if (h.Caption == field && (dim is null || h.DimensionUniqueName == dim)) return i;
        }
        return -1;
    }

    /// <summary>生成 pivotCacheDefinition 全文（核心段；hierarchies 待补）。</summary>
    public static string ToXml(PivotCacheInfo c)
    {
        var sb = new StringBuilder(1024);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<pivotCacheDefinition xmlns=\"{MainNs}\" " +
                  "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                  $"refreshedVersion=\"{c.RefreshedVersion}\" minRefreshableVersion=\"{c.MinRefreshableVersion}\" " +
                  $"createdVersion=\"{c.CreatedVersion}\" recordCount=\"0\"");
        if (c.RefreshedBy is not null) sb.Append($" refreshedBy=\"{Esc(c.RefreshedBy)}\"");
        if (c.RefreshedDate is { } dt && dt > 0)
            sb.Append($" refreshedDate=\"{dt.ToString("0.###############", CultureInfo.InvariantCulture)}\"");
        if (c.SourceType == 1) sb.Append(" saveData=\"0\" supportSubquery=\"1\" supportAdvancedDrill=\"1\"");
        sb.Append('>');

        if (c.SourceType == 1)
            sb.Append($"<cacheSource type=\"external\" connectionId=\"{c.ConnectionId}\"/>");
        else
            sb.Append("<cacheSource type=\"worksheet\"/>");

        sb.Append($"<cacheFields count=\"{c.Fields.Count}\">");
        foreach (var f in c.Fields)
        {
            sb.Append($"<cacheField name=\"{Esc(f.Name)}\"");
            if (f.Caption is not null) sb.Append($" caption=\"{Esc(f.Caption)}\"");
            sb.Append($" numFmtId=\"{f.NumFmtId}\" hierarchy=\"{f.Hierarchy}\" level=\"{f.Level}\"");
            if (f.SharedDates.Count > 0)
            {
                sb.Append("><sharedItems containsSemiMixedTypes=\"0\" containsNonDate=\"0\" containsDate=\"1\" " +
                          "containsString=\"0\"");
                sb.Append($" minDate=\"{FmtOle(f.SharedMin)}\" maxDate=\"{FmtOle(f.SharedMax)}\" count=\"{f.SharedDates.Count}\">");
                foreach (var v in f.SharedDates)
                    sb.Append($"<d v=\"{v}\"/>");
                sb.Append("</sharedItems><extLst><ext uri=\"{4F2E5C28-24EA-4eb8-9CBF-B6C8F9C3D259}\" " +
                          "xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\">" +
                          "<x15:cachedUniqueNames>");
                var baseName = f.Name.Substring(0, f.Name.LastIndexOf('.'));
                for (int di = 0; di < f.SharedDates.Count; di++)
                    sb.Append($"<x15:cachedUniqueName index=\"{di}\" name=\"{Esc(baseName)}.&amp;[{f.SharedDates[di]}]\"/>");
                sb.Append("</x15:cachedUniqueNames></ext></extLst></cacheField>");
            }
            else if (f.SharedStrings.Count > 0)
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

        if (c.Hierarchies.Count > 0)
        {
            sb.Append($"<cacheHierarchies count=\"{c.Hierarchies.Count}\">");
            for (int hi = 0; hi < c.Hierarchies.Count; hi++)
            {
                var h = c.Hierarchies[hi];
                sb.Append($"<cacheHierarchy uniqueName=\"{Esc(h.UniqueName)}\" caption=\"{Esc(h.Caption)}\"");
                if (h.Measure)
                {
                    sb.Append(" measure=\"1\"");
                    sb.Append($" displayFolder=\"{Esc(h.DisplayFolder ?? "")}\"");
                    if (h.MeasureGroup is not null) sb.Append($" measureGroup=\"{Esc(h.MeasureGroup)}\"");
                    sb.Append($" count=\"{h.Count}\"");
                    if (h.OneField) sb.Append(" oneField=\"1\"");
                    sb.Append(" hidden=\"1\"");
                    int aggCol = AggregatedColumn(c, h);
                    if (aggCol >= 0)
                        sb.Append($"><extLst><ext uri=\"{{B97F6D7D-B522-45F9-BDA1-12C45D357490}}\" " +
                                  "xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\">" +
                                  $"<x15:cacheHierarchy aggregatedColumn=\"{aggCol}\"/></ext></extLst></cacheHierarchy>");
                    else
                        sb.Append("/>");
                    continue;
                }
                if (h.Attribute) sb.Append(" attribute=\"1\"");
                if (h.MemberValueDatatype == 7) sb.Append(" time=\"1\"");
                if (h.DefaultMemberUniqueName is not null) sb.Append($" defaultMemberUniqueName=\"{Esc(h.DefaultMemberUniqueName)}\"");
                if (h.AllUniqueName is not null) sb.Append($" allUniqueName=\"{Esc(h.AllUniqueName)}\"");
                if (h.DimensionUniqueName is not null) sb.Append($" dimensionUniqueName=\"{Esc(h.DimensionUniqueName)}\"");
                sb.Append($" displayFolder=\"{Esc(h.DisplayFolder ?? "")}\"");
                sb.Append($" count=\"{h.Count}\" memberValueDatatype=\"{h.MemberValueDatatype}\" unbalanced=\"0\"");
                if (h.FieldsUsage.Count > 0)
                {
                    sb.Append($"><fieldsUsage count=\"{h.FieldsUsage.Count}\">");
                    foreach (var idx in h.FieldsUsage)
                        sb.Append($"<fieldUsage x=\"{idx}\"/>");
                    sb.Append("</fieldsUsage></cacheHierarchy>");
                }
                else
                {
                    sb.Append("/>");
                }
            }
            sb.Append("</cacheHierarchies>");
        }

        if (c.Dimensions.Count > 0)
        {
            sb.Append("<kpis count=\"0\"/>");
            sb.Append("<dimensions count=\"" + c.Dimensions.Count + "\">");
            foreach (var dim in c.Dimensions)
            {
                sb.Append("<dimension");
                if (dim.Measure) sb.Append(" measure=\"1\"");
                sb.Append($" name=\"{Esc(dim.Name)}\" uniqueName=\"{Esc(dim.UniqueName)}\" caption=\"{Esc(dim.Caption)}\"/>");
            }
            sb.Append("</dimensions>");
        }

        if (c.MeasureGroups.Count > 0)
        {
            sb.Append("<measureGroups count=\"" + c.MeasureGroups.Count + "\">");
            foreach (var mg in c.MeasureGroups)
                sb.Append($"<measureGroup name=\"{Esc(mg.Name)}\" caption=\"{Esc(mg.Caption)}\"/>");
            sb.Append("</measureGroups>");
        }

        if (c.Maps.Count > 0)
        {
            sb.Append("<maps count=\"" + c.Maps.Count + "\">");
            foreach (var (mg, dim) in c.Maps)
                sb.Append($"<map measureGroup=\"{mg}\" dimension=\"{dim}\"/>");
            sb.Append("</maps>");
        }

        if (c.SourceType == 1)
            sb.Append("<extLst><ext uri=\"{725AE2AE-9491-48be-B2B4-4EB974FC3084}\" " +
                      "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">" +
                      "<x14:pivotCacheDefinition supportSubqueryNonVisual=\"1\" supportSubqueryCalcMem=\"1\" " +
                      "supportAddCalcMems=\"1\"/></ext></extLst>");

        sb.Append("</pivotCacheDefinition>");
        return sb.ToString();
    }
}
