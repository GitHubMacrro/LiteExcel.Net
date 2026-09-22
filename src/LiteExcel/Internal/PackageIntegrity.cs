using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace LiteExcel.Internal;

/// <summary>
/// 保存后包结构自检：验证产出的 OPC 包不存在会致 Excel 报「已修复的记录」或闪退的结构损坏。
/// 检查项：
/// <list type="bullet">
/// <item><b>孤儿 .rels</b>：关系文件（*.rels）的父部件不存在（如删了 drawingN.xml 却漏删其 .rels）。</item>
/// <item><b>悬空引用</b>：某 .rels 的 Target 解析后指向不存在的部件。</item>
/// <item><b>悬空 CT Override</b>：[Content_Types].xml 声明的部件不存在。</item>
/// </list>
/// 目的：把「悄悄产出损坏文件 → 用户在 Excel 才发现闪退」提前到「保存时立即失败并指出位置」。
/// </summary>
internal static class PackageIntegrity
{
    /// <summary>一条结构违规记录。</summary>
    internal sealed class Violation
    {
        /// <summary>违规类型（Orphan-Rels / Dangling-Rel / CT-Override）。</summary>
        public string Kind = "";
        /// <summary>人类可读的定位描述。</summary>
        public string Detail = "";
        /// <inheritdoc/>
        public override string ToString() => $"[{Kind}] {Detail}";
    }

    /// <summary>校验一个已打开的 zip 包，返回全部结构违规（空列表表示通过）。</summary>
    public static List<Violation> Validate(ZipArchive zip)
    {
        var violations = new List<Violation>();
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (var e in zip.Entries) names.Add(e.FullName);

        // 1) [Content_Types].xml 的 Override 是否都指向存在的部件
        var ct = zip.GetEntry("[Content_Types].xml");
        if (ct is not null)
            foreach (var part in ParseCtOverrides(ct))
                if (!names.Contains(part))
                    violations.Add(new Violation { Kind = "CT-Override", Detail = $"Content_Types 声明了不存在的部件：/{part}" });

        // 2) 逐个 .rels：父部件是否存在（孤儿）+ Target 是否存在（悬空）
        foreach (var e in zip.Entries)
        {
            var name = e.FullName;
            if (!name.EndsWith(".rels", StringComparison.Ordinal)) continue;
            if (name == "_rels/.rels") continue; // 根关系无父部件

            int idx = name.LastIndexOf("_rels/", StringComparison.Ordinal);
            var dir = idx > 0 ? name.Substring(0, idx) : "";
            var file = name.Substring(idx + "_rels/".Length);
            if (file.EndsWith(".rels", StringComparison.Ordinal))
                file = file.Substring(0, file.Length - ".rels".Length);
            var parent = dir + file;
            if (!names.Contains(parent))
                violations.Add(new Violation { Kind = "Orphan-Rels", Detail = $"关系文件无对应部件：{name}（父部件 {parent} 不存在）" });

            string xml;
            using (var s = e.Open())
            using (var r = new StreamReader(s, Encoding.UTF8))
                xml = r.ReadToEnd();
            foreach (var rel in XlsxWriter.ParseRels(xml))
            {
                if (string.IsNullOrEmpty(rel.Target)) continue;
                if (rel.TargetMode.Equals("External", StringComparison.OrdinalIgnoreCase)) continue;
                if (rel.Target.IndexOf("://", StringComparison.Ordinal) >= 0) continue;
                var abs = ResolveTarget(dir, rel.Target);
                if (!names.Contains(abs))
                    violations.Add(new Violation { Kind = "Dangling-Rel", Detail = $"{name} → {rel.Target}（解析为 {abs}，部件不存在）" });
            }
        }
        return violations;
    }

    /// <summary>校验可寻址的包流；存在结构违规时抛出 <see cref="LiteExcelException"/>（阻止写出损坏文件）。</summary>
    public static void ValidateOrThrow(Stream packageStream)
    {
        packageStream.Position = 0;
        using var zip = new ZipArchive(packageStream, ZipArchiveMode.Read, leaveOpen: true);
        var v = Validate(zip);
        if (v.Count > 0)
            throw new LiteExcelException(
                "包结构自检失败（已阻止写出损坏文件，避免 Excel 打开时报修复/闪退）：\n - "
                + string.Join("\n - ", v.Select(x => x.ToString())));
    }

    private static string ResolveTarget(string baseDir, string target)
    {
        target = target.Replace('\\', '/');
        if (target.StartsWith("/", StringComparison.Ordinal)) return target.TrimStart('/');
        var combined = baseDir.Length == 0 ? target : baseDir + "/" + target;
        var segs = combined.Split('/');
        var stack = new List<string>(segs.Length);
        foreach (var seg in segs)
        {
            if (seg.Length == 0 || seg == ".") continue;
            if (seg == "..")
            {
                if (stack.Count > 0) stack.RemoveAt(stack.Count - 1);
                continue;
            }
            stack.Add(seg);
        }
        return string.Join("/", stack);
    }

    private static IEnumerable<string> ParseCtOverrides(ZipArchiveEntry ct)
    {
        string xml;
        using (var s = ct.Open())
        using (var r = new StreamReader(s, Encoding.UTF8))
            xml = r.ReadToEnd();
        XDocument doc;
        try { doc = XDocument.Parse(xml); }
        catch { yield break; }
        foreach (var el in doc.Descendants())
            if (el.Name.LocalName == "Override")
            {
                var pn = (string?)el.Attribute("PartName");
                if (!string.IsNullOrEmpty(pn)) yield return pn.TrimStart('/');
            }
    }
}
