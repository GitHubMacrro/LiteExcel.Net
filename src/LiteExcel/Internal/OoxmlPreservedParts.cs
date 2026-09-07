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
}
