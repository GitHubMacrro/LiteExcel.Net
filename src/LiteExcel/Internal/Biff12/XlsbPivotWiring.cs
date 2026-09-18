using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace LiteExcel.Internal.Biff12;

/// <summary>
/// 跨格式接线：把 xlsb 的 pivotCacheDefinition / pivotTable 部件转码为 OOXML，
/// 并合成 workbook 的 &lt;pivotCaches&gt; 与 cacheId 映射（Stage D 接线）。
///
/// cacheId 规则（真实样本逐字节标定）：
///   workbook.bin 中 0x0182/0x046D 记录按出现顺序构成缓存列表，位置 i 即 xlsx cacheId；
///   每条记录的 flags(u32) 即引用它的 pivotTable 的 idCache（BrtBeginSXView@28）；
///   记录的 rId 经 workbook.bin.rels 解析到 pivotCacheDefinitionN.bin 部件。
/// </summary>
internal static class XlsbPivotWiring
{
    internal sealed class CacheRef
    {
        public uint Flags;
        public string RelId = "";
        public string? Part;
    }

    /// <summary>解析 workbook.bin 的缓存引用记录（0x0182 / 0x046D），返回按顺序的缓存列表。</summary>
    public static List<CacheRef> ParseCacheRefs(byte[] workbookBin, IReadOnlyDictionary<string, string> relIdToTarget)
    {
        var list = new List<CacheRef>();
        foreach (var rec in Biff12Records.ReadAll(workbookBin))
        {
            string relId;
            uint flags;
            if (rec.Rt == 0x0182)
            {
                if (rec.Data.Length < 8) continue;
                flags = Biff12Records.ReadU32(rec.Data, 0);
                int o = 4;
                relId = ReadWideString(rec.Data, ref o);
            }
            else if (rec.Rt == 0x046D)
            {
                if (rec.Data.Length < 6) continue;
                flags = Biff12Records.ReadU32(rec.Data, 0);
                int cch = Biff12Records.ReadU16(rec.Data, 4);
                if (6 + cch * 2 > rec.Data.Length) continue;
                relId = Encoding.Unicode.GetString(rec.Data, 6, cch * 2);
            }
            else continue;

            var cache = new CacheRef { Flags = flags, RelId = relId };
            if (relId.Length > 0 && relIdToTarget.TryGetValue(relId, out var target))
                cache.Part = NormalizeCachePart(target);
            list.Add(cache);
        }
        return list;
    }

    private static string NormalizeCachePart(string target)
    {
        if (target.StartsWith("/", StringComparison.Ordinal)) return target.TrimStart('/');
        if (target.StartsWith("xl/", StringComparison.Ordinal)) return target;
        return "xl/" + target;
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

    /// <summary>由 workbook.bin.rels 文本构建 relId → target 映射。</summary>
    public static Dictionary<string, string> RelIdToTarget(string workbookRelsXml)
    {
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        if (string.IsNullOrEmpty(workbookRelsXml)) return map;
        try
        {
            var doc = XDocument.Parse(workbookRelsXml);
            var ns = doc.Root?.GetDefaultNamespace() ?? XNamespace.None;
            foreach (var el in doc.Root?.Elements(ns + "Relationship") ?? Enumerable.Empty<XElement>())
            {
                var id = (string?)el.Attribute("Id");
                var target = (string?)el.Attribute("Target");
                if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(target))
                    map[id] = target;
            }
        }
        catch { }
        return map;
    }

    /// <summary>合成 workbook &lt;pivotCaches&gt;：cacheId = 位置，r:id = 原 workbook.bin 关系 Id（写出时经 keptIdMap 重映射）。</summary>
    /// <summary>合成 workbook &lt;pivotCaches&gt;：仅包含被透视表引用的缓存（cacheId = 位置）。
    /// slicerData 缓存（无透视表引用）改由 <see cref="BuildPivotCachesExtLstXml"/> 写入 x14 extLst（真实 Excel 行为）。</summary>
    public static string BuildPivotCachesXml(IReadOnlyList<CacheRef> caches, IReadOnlyCollection<int> usedCacheIds)
    {
        var sb = new StringBuilder();
        sb.Append("<pivotCaches>");
        for (int i = 0; i < caches.Count; i++)
        {
            if (usedCacheIds.Contains(i)) sb.Append($"<pivotCache cacheId=\"{i}\" r:id=\"{caches[i].RelId}\"/>");
        }
        sb.Append("</pivotCaches>");
        return sb.ToString();
    }

    /// <summary>x14 extLst &lt;pivotCaches&gt;：包含未被透视表引用、但仍是数据模型缓存（如 slicerData）的项。</summary>
    public static string? BuildPivotCachesExtLstXml(IReadOnlyList<CacheRef> caches, IReadOnlyCollection<int> usedCacheIds)
    {
        var extra = new List<int>();
        for (int i = 0; i < caches.Count; i++)
            if (!usedCacheIds.Contains(i)) extra.Add(i);
        if (extra.Count == 0) return null;
        var sb = new StringBuilder();
        sb.Append("<ext uri=\"{876F7934-8845-4945-9796-88D515C7AA90}\" " +
                  "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:pivotCaches>");
        foreach (var i in extra)
            sb.Append($"<pivotCache cacheId=\"{i}\" r:id=\"{caches[i].RelId}\"/>");
        sb.Append("</x14:pivotCaches></ext>");
        return sb.ToString();
    }

    /// <summary>把 pivotTableN.xml 的 cacheId 改写为位置值（由该表的 idCache 经 flags→位置映射得到）。</summary>
    public static string PatchCacheId(string pivotTableXml, int cacheId)
    {
        int idx = pivotTableXml.IndexOf(" cacheId=\"", StringComparison.Ordinal);
        if (idx < 0) return pivotTableXml;
        int start = idx + " cacheId=\"".Length;
        int end = pivotTableXml.IndexOf('"', start);
        if (end < 0) return pivotTableXml;
        return pivotTableXml.Substring(0, start) + cacheId + pivotTableXml.Substring(end);
    }
}
