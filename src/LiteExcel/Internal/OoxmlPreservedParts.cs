using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace LiteExcel.Internal;

/// <summary>
/// 打开工作簿时捕获的、写入器不会重建的 OOXML 部件。
/// 保存时按二进制透传，避免未映射部件（宏/主题/绘图/图表/表格等）被静默删除。
/// </summary>
internal sealed class OoxmlPreservedParts
{
    /// <summary>按包路径保存的原始部件字节（不含写入器重建的条目与合并用的 rels） </summary>
    public readonly Dictionary<string, byte[]> Parts = new();

    /// <summary>需要合并的 rels 文件（根 / 工作簿 / 工作表），key 为包路径 </summary>
    public readonly Dictionary<string, string> Rels = new();

    /// <summary>原始 [Content_Types].xml 的 Default 声明 </summary>
    public readonly List<(string Extension, string ContentType)> DefaultTypes = new();

    /// <summary>原始 [Content_Types].xml 的 Override 声明 </summary>
    public readonly List<(string PartName, string ContentType)> OverrideTypes = new();

    /// <summary>工作簿宿主的 VBA 代码名（workbookPr@codeName），保存时写回重建的 workbook.xml，保持与保留的 vbaProject 绑定 </summary>
    public string? WorkbookCodeName { get; set; }

    /// <summary>打开时捕获的 workbook.xml 中 bookViews 元素原始 XML，保存时按 OOXML 顺序回写。</summary>
    public string? BookViewsXml { get; set; }

    /// <summary>打开时捕获的 workbook.xml 中 definedNames 元素原始 XML，保存时按 OOXML 顺序回写。</summary>
    public string? DefinedNamesXml { get; set; }

    /// <summary>打开时捕获的 workbook.xml 中 pivotCaches 元素原始 XML，保存时同步重映射关系 ID 后回写。</summary>
    public string? PivotCachesXml { get; set; }

    /// <summary>打开时捕获的 workbook.xml 中 externalReferences 元素原始 XML，保存时同步重映射关系 ID 后回写。</summary>
    public string? ExternalReferencesXml { get; set; }

    /// <summary>打开时捕获的 workbook.xml 中 extLst 元素原始 XML（含 x14/x15 slicerCaches / timelineCaches 等），
    /// 保存时按新 rel Id 重映射后回写（schema 位于 calcPr 之后、</workbook> 之前）。切片器/日程表缓存引用住在其中。</summary>
    public string? WorkbookExtLstXml { get; set; }

    /// <summary>XLSB verbatim 保留：原始 workbook.bin / styles.bin / sharedStrings.bin / sheetN.bin 字节。
    /// 当工作簿结构不变且无单元格修改时，XlsbWriter 原样写出这些字节而非重建，
    /// 保留透视表/切片器等 BIFF12 宿主记录。key = 包内路径（如 "xl/workbook.bin"）</summary>
    public Dictionary<string, byte[]>? VerbatimBinaries { get; set; }

    /// <summary>XLSX verbatim 保留：原始 styles.xml / sharedStrings.xml / sheetN.xml 字节。
    /// 当工作簿结构不变且无单元格修改时，XlsxWriter 原样写出这些字节而非重建，
    /// 保留 slicerStyles / timelineStyles / pivotButton XF 等扩展样式。
    /// key = 包内路径（如 "xl/styles.xml"）</summary>
    public Dictionary<string, byte[]>? VerbatimXmlParts { get; set; }

    /// <summary>捕获 zip 中写入器不重建的部件与 rels。sheetCount 用于排除所有工作表/批注重建条目。
    /// <paramref name="binary"/> = true 时按 xlsb 容器布局排除（.bin 工作表/工作簿/styles 等）</summary>
    public static OoxmlPreservedParts Capture(ZipArchive zip, int sheetCount, bool binary = false)
    {
        var preserved = new OoxmlPreservedParts();
        var rebuilt = BuildRebuiltEntries(sheetCount, binary);

        if (binary)
            preserved.VerbatimBinaries = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        else
            preserved.VerbatimXmlParts = new Dictionary<string, byte[]>(StringComparer.Ordinal);

        foreach (var entry in zip.Entries)
        {
            var name = entry.FullName;

            if (name == "[Content_Types].xml")
            {
                ParseContentTypes(entry, preserved);
                continue;
            }

            // 超级表已由 XlsxReader 读入 sheet.Tables 并带原始 XML 回写；从保留部件剔除，
            // 避免 TablePlan 重建（table13+）与保留部件透传（table1-12）双写导致表重复/保真丢失。
            if (IsTablePart(name))
                continue;

            // 捕获根、工作簿和工作表关系，供保存时合并。
            if (IsMergeRelsPath(name, binary))
            {
                preserved.Rels[name] = ReadText(entry);
                continue;
            }

            // 捕获写入器会重建的 XLSB 二进制部件。
            if (binary && rebuilt.Contains(name))
            {
                preserved.VerbatimBinaries![name] = ReadBytes(entry);
                continue;
            }

            // 捕获写入器会重建的 XLSX XML 部件。
            if (!binary && rebuilt.Contains(name))
            {
                preserved.VerbatimXmlParts![name] = ReadBytes(entry);
                continue;
            }

            if (rebuilt.Contains(name)) continue;

            preserved.Parts[name] = ReadBytes(entry);
        }

        // 兜底：若第一个 [Content_Types].xml 未遍历到（正常会遍历到），这里再解析一次
        var ct = zip.GetEntry("[Content_Types].xml");
        if (ct is not null && preserved.DefaultTypes.Count == 0 && preserved.OverrideTypes.Count == 0)
            ParseContentTypes(ct, preserved);

        return preserved;
    }

    /// <summary>是否为超级表定义部件（xl/tables/table{N}.xml）。注意不排除 xl/tables/_rels/*.rels：表 rels 引用 queryTable 等子部件，须透传保留。 </summary>
    private static bool IsTablePart(string name)
        => name.StartsWith("xl/tables/table", StringComparison.Ordinal)
            && name.EndsWith(".xml", StringComparison.Ordinal);

    /// <summary>写入器会整体重建（不保留）的包条目 </summary>
    internal static HashSet<string> BuildRebuiltEntries(int sheetCount, bool binary = false)
    {
        if (binary)
        {
            var setB = new HashSet<string>(StringComparer.Ordinal)
            {
                "[Content_Types].xml",
                "_rels/.rels",
                "xl/workbook.bin",
                "xl/_rels/workbook.bin.rels",
                "xl/sharedStrings.bin",
                "xl/styles.bin",
                "xl/calcChain.bin",   // 陈旧计算链不透传，由 Excel 在打开时重建。
                "docProps/core.xml",
                "docProps/app.xml",
            };
            for (int i = 1; i <= sheetCount; i++)
            {
                setB.Add($"xl/worksheets/sheet{i}.bin");
                setB.Add($"xl/worksheets/_rels/sheet{i}.bin.rels");
            }
            return setB;
        }

        var set = new HashSet<string>(StringComparer.Ordinal)
        {
            "[Content_Types].xml",
            "_rels/.rels",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/sharedStrings.xml",
            "xl/styles.xml",
            "xl/calcChain.xml",  // 陈旧计算链不透传，由 Excel 在打开时重建。
            "docProps/core.xml",
            "docProps/app.xml",
        };
        for (int i = 1; i <= sheetCount; i++)
        {
            set.Add($"xl/worksheets/sheet{i}.xml");
            set.Add($"xl/worksheets/_rels/sheet{i}.xml.rels");
            set.Add($"xl/comments{i}.xml");
        }
        return set;
    }

    private static bool IsMergeRelsPath(string name, bool binary)
    {
        if (name == "_rels/.rels") return true;
        if (name == "xl/_rels/workbook.xml.rels" || name == "xl/_rels/workbook.bin.rels") return true;
        if (name.StartsWith("xl/worksheets/_rels/", StringComparison.Ordinal) && name.EndsWith(".rels", StringComparison.Ordinal))
            return true;
        return false;
    }

    private static void ParseContentTypes(ZipArchiveEntry entry, OoxmlPreservedParts preserved)
    {
        using var stream = entry.Open();
        var doc = XDocument.Load(stream);
        if (doc.Root is null) return;
        var ns = doc.Root.GetDefaultNamespace();

        foreach (var el in doc.Root.Elements(ns + "Default"))
        {
            var ext = (string?)el.Attribute("Extension") ?? "";
            var ct = (string?)el.Attribute("ContentType") ?? "";
            if (ext.Length > 0 && ct.Length > 0) preserved.DefaultTypes.Add((ext, ct));
        }
        foreach (var el in doc.Root.Elements(ns + "Override"))
        {
            var part = (string?)el.Attribute("PartName") ?? "";
            var ct = (string?)el.Attribute("ContentType") ?? "";
            if (part.Length > 0 && ct.Length > 0) preserved.OverrideTypes.Add((part, ct));
        }
    }

    private static string ReadText(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return reader.ReadToEnd();
    }

    private static byte[] ReadBytes(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return ms.ToArray();
    }

    /// <summary>
    /// 把从 xlsb 源捕获的保留部件转换为可用于 xlsx/xlsm 写出的形态（跨格式转换 Stage A）。
    /// 仅保留「格式无关」部件（两种容器中同构），其余 BIFF12 (.bin) 记录不混入 OOXML 包；
    /// 同时把 .bin 关系路径/目标与内容类型声明按目标容器重写，并剔除指向未写出部件的关系，
    /// 避免产生悬空引用导致 Excel 修复。
    /// </summary>
    internal OoxmlPreservedParts ToXlsxCompatible()
    {
        var result = new OoxmlPreservedParts
        {
            // 非 null 标记：XlsxWriter 以此判断保留部件可用（不再整体丢弃）。
            VerbatimXmlParts = new Dictionary<string, byte[]>(StringComparer.Ordinal),
            WorkbookCodeName = WorkbookCodeName,
            BookViewsXml = BookViewsXml,
            DefinedNamesXml = DefinedNamesXml,
        };

        foreach (var kv in Parts)
            if (IsFormatAgnosticForXlsx(kv.Key))
                result.Parts[kv.Key] = kv.Value;

        // 跨格式转码（Stage C）：connections / queryTables / 数据模型。
        // 关系目标需从 .bin 重写为 .xml，故先记录映射再过滤 rels。
        var targetMap = new Dictionary<string, string>(StringComparer.Ordinal);
        if (Parts.TryGetValue("xl/connections.bin", out var connBin))
        {
            var connXml = Biff12.XlsbConnectionTranscoder.ToXml(Biff12.XlsbConnectionTranscoder.Parse(connBin));
            result.Parts["xl/connections.xml"] = Encoding.UTF8.GetBytes(connXml);
            targetMap["xl/connections.bin"] = "xl/connections.xml";
            result.OverrideTypes.Add(("/xl/connections.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"));
            // 注：type-102 链接表所需的 `_xlcn.LinkedTable_*` 定义名由 Workbook.BuildXlsbDefinedNamesXml 统一合成。
        }
        if (Parts.TryGetValue("xl/model/item.data", out var modelData))
        {
            result.Parts["xl/model/item.data"] = modelData;
            result.OverrideTypes.Add(("/xl/model/item.data", "application/vnd.openxmlformats-officedocument.model+data"));
        }

        // 绘图：xlsb 的 graphicFrame(compatSp) → xlsx 的 sp（否则 Excel 拒开）；drawing rels 原样保留。
        foreach (var kv in Parts)
        {
            if (kv.Key.StartsWith("xl/drawings/drawing", StringComparison.Ordinal) && kv.Key.EndsWith(".xml", StringComparison.Ordinal))
                result.Parts[kv.Key] = Encoding.UTF8.GetBytes(Biff12.XlsbDrawingTranscoder.Transcode(Encoding.UTF8.GetString(kv.Value)));
            else if (kv.Key.StartsWith("xl/drawings/_rels/drawing", StringComparison.Ordinal) && kv.Key.EndsWith(".rels", StringComparison.Ordinal))
                result.Parts[kv.Key] = kv.Value;
        }

        // 透视表（Stage D 接线）+ 切片器（Stage E 接线）：二者必须一起启用（透视表强依赖切片器）。
        // 注意：当前转码尚未覆盖全部变体（如 pivotFields 翻倍怪癖、formats/extLst、非 OLAP 切片器），
        // 启用前须确保 Excel 能无修复打开；暂以环境变量门控（默认关闭）。
        if (Environment.GetEnvironmentVariable("LITEXCEL_ENABLE_PIVOT_WIRING") == "1")
        {
            TranscodePivotParts(result, targetMap);
            TranscodeSlicerParts(result, targetMap);
        }
        foreach (var kv in Parts)
        {
            if (!kv.Key.StartsWith("xl/queryTables/queryTable", StringComparison.Ordinal)) continue;
            if (!kv.Key.EndsWith(".bin", StringComparison.Ordinal)) continue;
            int slash = kv.Key.LastIndexOf('/');
            var num = kv.Key.Substring(slash + 1); // queryTableN.bin
            var xmlPath = "xl/queryTables/" + num.Substring(0, num.Length - 4) + ".xml";
            result.Parts[xmlPath] = Encoding.UTF8.GetBytes(Biff12.XlsbConnectionTranscoder.ToQueryTableXml(kv.Value));
            targetMap[kv.Key] = xmlPath;
            result.OverrideTypes.Add(("/" + xmlPath, "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"));
        }

        // 内容类型：丢弃 xlsb 专有 Default（bin/data），Override 仅保留已复制且非 .bin 的部件。
        foreach (var (ext, ct) in DefaultTypes)
            if (!string.Equals(ext, "bin", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(ext, "data", StringComparison.OrdinalIgnoreCase))
                result.DefaultTypes.Add((ext, ct));
        foreach (var (part, ct) in OverrideTypes)
        {
            var path = part.TrimStart('/');
            // 仅保留实际已复制部件的 Override（含 vbaProject.bin 等 .bin 格式无关部件）。
            if (!result.Parts.ContainsKey(path)) continue;
            result.OverrideTypes.Add((part, ct));
        }

        // 关系：.bin.rels → .xml.rels；目标 .bin → .xml；剔除指向未写出部件的关系。
        foreach (var kv in Rels)
        {
            var newKey = RenameBinaryRelsKey(kv.Key);
            var baseDir = RelsBaseDir(newKey);
            result.Rels[newKey] = FilterRelsToAvailableParts(kv.Value, baseDir, result.Parts, targetMap);
        }

        return result;
    }

    /// <summary>
    /// 跨格式可直通的格式无关部件（xlsb / xlsx 中同构）。
    /// 注意：`xl/drawings/drawing*.xml` 是 xlsb 专有的 ActiveX 图形表达（xdr:graphicFrame + com14:compatSp），
    /// 混入 xlsx 会被 Excel 拒绝（0x800A03EC）；须由后续阶段的 drawing 转码处理，此处排除。
    /// 仅保留 `xl/drawings/vmlDrawing*`。
    /// </summary>
    private void TranscodePivotParts(OoxmlPreservedParts result, Dictionary<string, string> targetMap)
    {
        bool hasCache = false, hasTable = false;
        foreach (var kv in Parts)
        {
            if (kv.Key.StartsWith("xl/pivotCache/pivotCacheDefinition", StringComparison.Ordinal) && kv.Key.EndsWith(".bin", StringComparison.Ordinal))
            {
                var xmlPath = kv.Key.Substring(0, kv.Key.Length - 4) + ".xml";
                var xml = Biff12.XlsbPivotCacheTranscoder.ToXml(Biff12.XlsbPivotCacheTranscoder.Parse(kv.Value));
                result.Parts[xmlPath] = Encoding.UTF8.GetBytes(xml);
                targetMap[kv.Key] = xmlPath;
                result.OverrideTypes.Add(("/" + xmlPath, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"));
                hasCache = true;
            }
        }
        if (!hasCache) return;

        // cacheId 映射：workbook.bin 的 0x0182/0x046D 顺序 = xlsx cacheId；flags = pivotTable idCache。
        List<Biff12.XlsbPivotWiring.CacheRef> caches = new();
        if (VerbatimBinaries is not null && VerbatimBinaries.TryGetValue("xl/workbook.bin", out var wbBin))
        {
            string relsXml = Rels.TryGetValue("xl/_rels/workbook.bin.rels", out var rx) ? rx : "";
            var relMap = Biff12.XlsbPivotWiring.RelIdToTarget(relsXml);
            caches = Biff12.XlsbPivotWiring.ParseCacheRefs(wbBin, relMap);
        }
        var cacheIdByFlags = new Dictionary<uint, int>();
        for (int i = 0; i < caches.Count; i++)
            if (!cacheIdByFlags.ContainsKey(caches[i].Flags))
                cacheIdByFlags[caches[i].Flags] = i;

        var usedCacheIds = new HashSet<int>();
        foreach (var kv in Parts)
        {
            if (!kv.Key.StartsWith("xl/pivotTables/pivotTable", StringComparison.Ordinal) || !kv.Key.EndsWith(".bin", StringComparison.Ordinal))
                continue;
            var info = Biff12.XlsbPivotTableTranscoder.Parse(kv.Value);
            int cacheId = cacheIdByFlags.TryGetValue(info.CacheId, out var pos) ? pos : (int)info.CacheId;
            usedCacheIds.Add(cacheId);
            var xmlPath = kv.Key.Substring(0, kv.Key.Length - 4) + ".xml";
            var xml = Biff12.XlsbPivotWiring.PatchCacheId(Biff12.XlsbPivotTableTranscoder.ToXml(info), cacheId);
            result.Parts[xmlPath] = Encoding.UTF8.GetBytes(xml);
            targetMap[kv.Key] = xmlPath;
            result.OverrideTypes.Add(("/" + xmlPath, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"));
            hasTable = true;
        }

        // pivotTable 自身的 rels（指向 pivotCacheDefinition）：.bin.rels → .xml.rels，目标 .bin → .xml。
        foreach (var kv in Parts)
        {
            if (!kv.Key.StartsWith("xl/pivotTables/_rels/pivotTable", StringComparison.Ordinal) || !kv.Key.EndsWith(".bin.rels", StringComparison.Ordinal))
                continue;
            var newKey = kv.Key.Substring(0, kv.Key.Length - ".bin.rels".Length) + ".xml.rels";
            result.Parts[newKey] = Encoding.UTF8.GetBytes(FilterRelsToAvailableParts(
                Encoding.UTF8.GetString(kv.Value), RelsBaseDir(newKey), result.Parts, targetMap));
        }

        if (hasTable && caches.Count > 0)
        {
            result.PivotCachesXml = Biff12.XlsbPivotWiring.BuildPivotCachesXml(caches, usedCacheIds);
            _pivotCachesExt = Biff12.XlsbPivotWiring.BuildPivotCachesExtLstXml(caches, usedCacheIds);
        }
    }

    private string? _pivotCachesExt;

    /// <summary>Stage E 接线：slicerCacheN.bin / slicerN.bin → .xml + CT Override + workbook extLst + sheet rels 目标重写。</summary>
    private void TranscodeSlicerParts(OoxmlPreservedParts result, Dictionary<string, string> targetMap)
    {
        var slicerCacheRels = new List<(string RelId, string XmlPath)>();
        bool any = false;
        foreach (var kv in Parts)
        {
            if (kv.Key.StartsWith("xl/slicerCaches/slicerCache", StringComparison.Ordinal) && kv.Key.EndsWith(".bin", StringComparison.Ordinal))
            {
                var xmlPath = kv.Key.Substring(0, kv.Key.Length - 4) + ".xml";
                var info = Biff12.XlsbSlicerTranscoder.ParseCache(kv.Value);
                result.Parts[xmlPath] = Encoding.UTF8.GetBytes(Biff12.XlsbSlicerTranscoder.ToSlicerCacheXml(info));
                targetMap[kv.Key] = xmlPath;
                result.OverrideTypes.Add(("/" + xmlPath, "application/vnd.ms-excel.slicerCache+xml"));
                any = true;
            }
            else if (kv.Key.StartsWith("xl/slicers/slicer", StringComparison.Ordinal) && kv.Key.EndsWith(".bin", StringComparison.Ordinal))
            {
                var xmlPath = kv.Key.Substring(0, kv.Key.Length - 4) + ".xml";
                var list = Biff12.XlsbSlicerTranscoder.ParseSlicers(kv.Value);
                result.Parts[xmlPath] = Encoding.UTF8.GetBytes(Biff12.XlsbSlicerTranscoder.ToSlicerXml(list));
                targetMap[kv.Key] = xmlPath;
                result.OverrideTypes.Add(("/" + xmlPath, "application/vnd.ms-excel.slicer+xml"));
                any = true;
            }
        }
        if (!any) return;

        // workbook extLst x14 slicerCaches：由 workbook.bin.rels 中 slicerCache 关系顺序合成。
        if (Rels.TryGetValue("xl/_rels/workbook.bin.rels", out var wbRels))
        {
            foreach (var m in System.Text.RegularExpressions.Regex.Matches(wbRels,
                "<Relationship Id=\"([^\"]+)\"[^>]*/slicerCache\"[^>]*Target=\"([^\"]+)\""))
            {
                var id = ((System.Text.RegularExpressions.Match)m).Groups[1].Value;
                var target = ((System.Text.RegularExpressions.Match)m).Groups[2].Value;
                var part = target.StartsWith("xl/", StringComparison.Ordinal) ? target : "xl/" + target.TrimStart('/');
                var xmlPath = part.EndsWith(".bin", StringComparison.Ordinal) ? part.Substring(0, part.Length - 4) + ".xml" : part;
                slicerCacheRels.Add((id, xmlPath));
            }
        }
        if (slicerCacheRels.Count > 0 || _pivotCachesExt is not null)
        {
            var sb = new StringBuilder();
            sb.Append("<extLst>");
            if (_pivotCachesExt is not null) sb.Append(_pivotCachesExt);
            if (slicerCacheRels.Count > 0)
            {
                sb.Append("<ext uri=\"{BBE1A952-AA13-448e-AADC-164F8A28A991}\" " +
                          "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:slicerCaches>");
                foreach (var (id, _) in slicerCacheRels)
                    sb.Append($"<x14:slicerCache r:id=\"{id}\"/>");
                sb.Append("</x14:slicerCaches></ext>");
            }
            sb.Append("</extLst>");
            result.WorkbookExtLstXml = sb.ToString();
        }

        // sheet rels：slicer 关系目标 .bin → .xml。
        foreach (var kv in Rels)
        {
            if (!kv.Key.StartsWith("xl/worksheets/_rels/", StringComparison.Ordinal)) continue;
            if (kv.Value.IndexOf("/slicer\"", StringComparison.Ordinal) < 0) continue;
            var newKey = RenameBinaryRelsKey(kv.Key);
            result.Rels[newKey] = FilterRelsToAvailableParts(kv.Value, RelsBaseDir(newKey), result.Parts, targetMap);
        }
    }

    /// <summary>
    /// 跨格式可直通的格式无关部件（xlsb / xlsx 中同构）。
    /// 注意：`xl/drawings/drawing*.xml` 是 xlsb 专有的 ActiveX 图形表达（xdr:graphicFrame + com14:compatSp），
    /// 混入 xlsx 会被 Excel 拒绝（0x800A03EC）；须由后续阶段的 drawing 转码处理，此处排除。
    /// 仅保留 `xl/drawings/vmlDrawing*`。
    /// </summary>
    private static bool IsFormatAgnosticForXlsx(string path)
    {
        if (path.StartsWith("customXml/", StringComparison.Ordinal)) return true;
        if (path == "docProps/custom.xml") return true;
        if (path == "xl/theme/theme1.xml") return true;
        if (path == "xl/vbaProject.bin") return true;
        if (path.StartsWith("xl/media/", StringComparison.Ordinal)) return true;
        if (path.StartsWith("xl/drawings/", StringComparison.Ordinal))
            return path.IndexOf("vmlDrawing", StringComparison.OrdinalIgnoreCase) >= 0;
        if (path.StartsWith("xl/activeX/", StringComparison.Ordinal)) return true;
        if (path.StartsWith("xl/printerSettings/", StringComparison.Ordinal)) return true;
        if (path.StartsWith("xl/ctrlProps/", StringComparison.Ordinal)) return true;
        return false;
    }

    private static string RenameBinaryRelsKey(string path)
    {
        if (path == "xl/_rels/workbook.bin.rels") return "xl/_rels/workbook.xml.rels";
        if (path.StartsWith("xl/worksheets/_rels/sheet", StringComparison.Ordinal)
            && path.EndsWith(".bin.rels", StringComparison.Ordinal))
            return path.Substring(0, path.Length - ".bin.rels".Length) + ".xml.rels";
        return path;
    }

    private static string RelsBaseDir(string relsPath)
    {
        // "xl/_rels/workbook.xml.rels" -> "xl"; "xl/worksheets/_rels/sheet1.xml.rels" -> "xl/worksheets"; "_rels/.rels" -> ""
        var idx = relsPath.IndexOf("/_rels/", StringComparison.Ordinal);
        return idx < 0 ? "" : relsPath.Substring(0, idx);
    }

    private static string ResolveRelTarget(string baseDir, string target)
    {
        target = target.Replace('\\', '/');
        if (target.StartsWith("/", StringComparison.Ordinal)) return target.TrimStart('/');
        var combined = (string.IsNullOrEmpty(baseDir) ? "" : baseDir + "/") + target;
        var stack = new List<string>();
        foreach (var p in combined.Split('/'))
        {
            if (string.IsNullOrEmpty(p) || p == ".") continue;
            if (p == "..") { if (stack.Count > 0) stack.RemoveAt(stack.Count - 1); continue; }
            stack.Add(p);
        }
        return string.Join("/", stack);
    }

    /// <summary>剔除指向未写出部件的关系（保留 External），并对 .bin→.xml 转码部件重写目标，避免悬空引用。</summary>
    private static string FilterRelsToAvailableParts(string relsXml, string baseDir, IReadOnlyDictionary<string, byte[]> available,
        IReadOnlyDictionary<string, string>? targetMap = null)
    {
        if (string.IsNullOrEmpty(relsXml)) return relsXml;
        XDocument doc;
        try { doc = XDocument.Parse(relsXml); }
        catch { return relsXml; }
        var root = doc.Root;
        if (root is null) return relsXml;
        var ns = root.GetDefaultNamespace();
        var keep = new List<XElement>();
        foreach (var el in root.Elements(ns + "Relationship"))
        {
            var targetMode = (string?)el.Attribute("TargetMode") ?? "";
            if (string.Equals(targetMode, "External", StringComparison.OrdinalIgnoreCase)) { keep.Add(el); continue; }
            var target = (string?)el.Attribute("Target") ?? "";
            var abs = ResolveRelTarget(baseDir, target);
            if (targetMap is not null && targetMap.TryGetValue(abs, out var mapped))
            {
                // 同一目录、仅扩展名不同（connections.bin→xml、queryTables/queryTableN.bin→xml）。
                el.SetAttributeValue("Target", target.Substring(0, target.Length - 4) + ".xml");
                abs = mapped;
            }
            if (available.ContainsKey(abs)) keep.Add(el);
        }
        root.ReplaceNodes(keep);
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + doc.ToString(SaveOptions.DisableFormatting);
    }
}
