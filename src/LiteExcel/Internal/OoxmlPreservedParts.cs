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

    /// <summary>捕获 zip 中写入器不重建的部件与 rels。sheetCount 用于排除所有工作表/批注重建条目 </summary>
    public static OoxmlPreservedParts Capture(ZipArchive zip, int sheetCount)
    {
        var preserved = new OoxmlPreservedParts();
        var rebuilt = BuildRebuiltEntries(sheetCount);

        foreach (var entry in zip.Entries)
        {
            var name = entry.FullName;

            if (name == "[Content_Types].xml")
            {
                ParseContentTypes(entry, preserved);
                continue;
            }

            // 需要合并的 rels（含重建的根/工作簿/工作表 rels）先捕获，供保存时合并
            if (IsMergeRelsPath(name))
            {
                preserved.Rels[name] = ReadText(entry);
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
    internal static HashSet<string> BuildRebuiltEntries(int sheetCount)
    {
        var set = new HashSet<string>(StringComparer.Ordinal)
        {
            "[Content_Types].xml",
            "_rels/.rels",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/sharedStrings.xml",
            "xl/styles.xml",
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

    private static bool IsMergeRelsPath(string name)
    {
        if (name == "_rels/.rels") return true;
        if (name == "xl/_rels/workbook.xml.rels") return true;
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
