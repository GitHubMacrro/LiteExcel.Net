using LiteExcel;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Xunit;

namespace LiteExcel.Tests;

/// <summary>
/// P0-26~P0-29 引用完整性：open-save 后包内所有 rel 目标必须存在，
/// 且 sheet XML / workbook.xml 内引用的 rId 必须能在对应 rels 中解析。
/// 旧断言只查部件 blob 是否存在，无法发现引用被丢弃或 rId 重编号后悬空。
/// </summary>
public class ReferenceIntegrityTests
{
    private const string PkgRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string OfficeRelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static string GetTempFile(string ext = ".xlsx") =>
        Path.Combine(Path.GetTempPath(), $"litexlsx_refint_{Guid.NewGuid():N}{ext}");

    private static string ReadText(ZipArchive zip, string entry)
    {
        var e = zip.GetEntry(entry);
        if (e is null) return "";
        using var s = e.Open();
        using var r = new StreamReader(s, Encoding.UTF8);
        return r.ReadToEnd();
    }

    private static void WriteEntry(ZipArchive zip, string entry, string content)
    {
        var e = zip.GetEntry(entry) ?? zip.CreateEntry(entry);
        using var s = e.Open();
        s.SetLength(0);
        var bytes = Encoding.UTF8.GetBytes(content);
        s.Write(bytes, 0, bytes.Length);
    }

    private static void AddContentTypeOverride(ZipArchive zip, string partName, string contentType)
    {
        var ct = ReadText(zip, "[Content_Types].xml");
        var doc = XDocument.Parse(ct);
        var ns = doc.Root!.GetDefaultNamespace();
        doc.Root.Add(new XElement(ns + "Override",
            new XAttribute("PartName", partName),
            new XAttribute("ContentType", contentType)));
        WriteEntry(zip, "[Content_Types].xml", doc.Declaration + doc.ToString(SaveOptions.DisableFormatting));
    }

    private static void AddWorkbookRel(ZipArchive zip, string id, string type, string target)
    {
        var rels = ReadText(zip, "xl/_rels/workbook.xml.rels");
        var doc = XDocument.Parse(rels);
        var ns = doc.Root!.GetDefaultNamespace();
        doc.Root.Add(new XElement(ns + "Relationship",
            new XAttribute("Id", id),
            new XAttribute("Type", type),
            new XAttribute("Target", target)));
        WriteEntry(zip, "xl/_rels/workbook.xml.rels", doc.Declaration + doc.ToString(SaveOptions.DisableFormatting));
    }

    /// <summary>把 XML 片段插到 workbook.xml 的指定锚点之后 </summary>
    private static void InsertIntoWorkbookXml(ZipArchive zip, string afterMarker, string fragment)
    {
        var wb = ReadText(zip, "xl/workbook.xml");
        int i = wb.IndexOf(afterMarker, StringComparison.Ordinal);
        Assert.True(i >= 0, $"workbook.xml 缺少锚点 {afterMarker}: {wb}");
        i += afterMarker.Length;
        WriteEntry(zip, "xl/workbook.xml", wb.Substring(0, i) + fragment + wb.Substring(i));
    }

    private static string ResolveRelTarget(string relsPath, string target)
    {
        if (target.StartsWith("/", StringComparison.Ordinal)) return target.TrimStart('/');
        int marker = relsPath.IndexOf("_rels/", StringComparison.Ordinal);
        string dir = marker < 0 ? "" : relsPath.Substring(0, marker);
        var stack = new List<string>(dir.TrimEnd('/').Split('/', StringSplitOptions.RemoveEmptyEntries));
        foreach (var seg in target.Replace('\\', '/').Split('/'))
        {
            if (seg == "..") { if (stack.Count > 0) stack.RemoveAt(stack.Count - 1); }
            else if (seg.Length > 0 && seg != ".") stack.Add(seg);
        }
        return string.Join("/", stack);
    }

    /// <summary>包内所有 rels 的非 External 目标都必须存在 </summary>
    private static void AssertNoDanglingRels(ZipArchive zip)
    {
        var names = new HashSet<string>(zip.Entries.Select(x => x.FullName), StringComparer.Ordinal);
        var dangling = new List<string>();
        foreach (var e in zip.Entries.Where(x => x.FullName.EndsWith(".rels", StringComparison.Ordinal)).ToList())
        {
            var doc = XDocument.Parse(ReadText(zip, e.FullName));
            var ns = doc.Root!.GetDefaultNamespace();
            foreach (var rel in doc.Root.Elements(ns + "Relationship"))
            {
                if ((string?)rel.Attribute("TargetMode") == "External") continue;
                var target = (string?)rel.Attribute("Target") ?? "";
                var abs = ResolveRelTarget(e.FullName, target);
                if (!names.Contains(abs)) dangling.Add($"{e.FullName}#{(string?)rel.Attribute("Id")} -> {abs}");
            }
        }
        Assert.True(dangling.Count == 0, "存在悬空 rel: " + string.Join(" | ", dangling));
    }

    /// <summary>某个 XML 部件内引用的所有 r:id 都必须能在其 rels 中解析 </summary>
    private static void AssertRelIdsResolvable(ZipArchive zip, string partPath, string relsPath)
    {
        var part = ReadText(zip, partPath);
        var rels = ReadText(zip, relsPath);
        var ids = new HashSet<string>(StringComparer.Ordinal);
        if (rels.Length > 0)
        {
            var doc = XDocument.Parse(rels);
            var ns = doc.Root!.GetDefaultNamespace();
            foreach (var rel in doc.Root.Elements(ns + "Relationship"))
                ids.Add((string?)rel.Attribute("Id") ?? "");
        }
        var bad = Regex.Matches(part, "r:id=\"([^\"]*)\"")
            .Select(m => m.Groups[1].Value)
            .Where(x => !ids.Contains(x))
            .Distinct(StringComparer.Ordinal)
            .ToList();
        Assert.True(bad.Count == 0, $"{partPath} 内 r:id 悬空: {string.Join(",", bad)}");
    }

    private static string RelTargetFor(ZipArchive zip, string relsPath, string relId)
    {
        var doc = XDocument.Parse(ReadText(zip, relsPath));
        var ns = doc.Root!.GetDefaultNamespace();
        foreach (var rel in doc.Root.Elements(ns + "Relationship"))
            if ((string?)rel.Attribute("Id") == relId) return (string?)rel.Attribute("Target") ?? "";
        return "";
    }

    private static string RelTypeFor(ZipArchive zip, string relsPath, string relId)
    {
        var doc = XDocument.Parse(ReadText(zip, relsPath));
        var ns = doc.Root!.GetDefaultNamespace();
        foreach (var rel in doc.Root.Elements(ns + "Relationship"))
            if ((string?)rel.Attribute("Id") == relId) return (string?)rel.Attribute("Type") ?? "";
        return "";
    }

    [Fact]
    public void OpenSave_PreservesSheetDrawingReference()
    {
        // P0-26: 保留的 drawing（图表/形状）没有对应新图片时，
        // 旧实现不写 <drawing>，sheet rels 的 drawing 关系变悬空 → Excel 认为无绘图，图表消失
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets["Sheet1"].SetValue("A1", "v");
            wb.SaveAs(file);

            using (var zip = new ZipArchive(File.Open(file, FileMode.Open, FileAccess.ReadWrite), ZipArchiveMode.Update))
            {
                WriteEntry(zip, "xl/drawings/drawing1.xml",
                    "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>");
                WriteEntry(zip, "xl/worksheets/_rels/sheet1.xml.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    $"<Relationships xmlns=\"{PkgRelNs}\">" +
                    $"<Relationship Id=\"rId1\" Type=\"{OfficeRelNs}/drawing\" Target=\"../drawings/drawing1.xml\"/>" +
                    "</Relationships>");
                var sheet = ReadText(zip, "xl/worksheets/sheet1.xml");
                WriteEntry(zip, "xl/worksheets/sheet1.xml",
                    sheet.Replace("</worksheet>", "<drawing r:id=\"rId1\"/></worksheet>"));
                AddContentTypeOverride(zip, "/xl/drawings/drawing1.xml",
                    "application/vnd.openxmlformats-officedocument.drawing+xml");
            }

            var opened = Excel.Open(file);
            opened.Worksheets[0].SetValue("B1", "x");
            opened.Save();

            using var check = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            var savedSheet = ReadText(check, "xl/worksheets/sheet1.xml");
            var m = Regex.Match(savedSheet, "<drawing r:id=\"([^\"]+)\"\\s*/>");
            Assert.True(m.Success, "sheet1.xml 丢失 <drawing> 元素: " + savedSheet);

            // 引用必须指向真实存在的 drawing 部件
            var relId = m.Groups[1].Value;
            Assert.Contains("/drawing", RelTypeFor(check, "xl/worksheets/_rels/sheet1.xml.rels", relId));
            Assert.Contains("drawing1.xml", RelTargetFor(check, "xl/worksheets/_rels/sheet1.xml.rels", relId));
            Assert.NotNull(check.GetEntry("xl/drawings/drawing1.xml"));

            AssertRelIdsResolvable(check, "xl/worksheets/sheet1.xml", "xl/worksheets/_rels/sheet1.xml.rels");
            AssertNoDanglingRels(check);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void OpenSave_PreservesSheetDrawingReference_WhenHyperlinkShiftsRelIds()
    {
        // P0-26 边界：保留 rel 被 MergeRelsXml 重新编号时，<drawing r:id> 必须同步改写
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets["Sheet1"].SetValue("A1", "v");
            wb.SaveAs(file);

            using (var zip = new ZipArchive(File.Open(file, FileMode.Open, FileAccess.ReadWrite), ZipArchiveMode.Update))
            {
                WriteEntry(zip, "xl/drawings/drawing1.xml",
                    "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>");
                // drawing rel 排在外部超链接之后，重新编号后 Id 会变化
                WriteEntry(zip, "xl/worksheets/_rels/sheet1.xml.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    $"<Relationships xmlns=\"{PkgRelNs}\">" +
                    $"<Relationship Id=\"rId7\" Type=\"{OfficeRelNs}/hyperlink\" Target=\"https://example.com\" TargetMode=\"External\"/>" +
                    $"<Relationship Id=\"rId9\" Type=\"{OfficeRelNs}/drawing\" Target=\"../drawings/drawing1.xml\"/>" +
                    "</Relationships>");
                var sheet = ReadText(zip, "xl/worksheets/sheet1.xml");
                WriteEntry(zip, "xl/worksheets/sheet1.xml",
                    sheet.Replace("</worksheet>", "<drawing r:id=\"rId9\"/></worksheet>"));
                AddContentTypeOverride(zip, "/xl/drawings/drawing1.xml",
                    "application/vnd.openxmlformats-officedocument.drawing+xml");
            }

            var opened = Excel.Open(file);
            opened.Worksheets[0].SetValue("B1", "x");
            opened.Save();

            using var check = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            var savedSheet = ReadText(check, "xl/worksheets/sheet1.xml");
            var m = Regex.Match(savedSheet, "<drawing r:id=\"([^\"]+)\"\\s*/>");
            Assert.True(m.Success, "sheet1.xml 丢失 <drawing> 元素: " + savedSheet);
            Assert.Contains("/drawing", RelTypeFor(check, "xl/worksheets/_rels/sheet1.xml.rels", m.Groups[1].Value));

            AssertRelIdsResolvable(check, "xl/worksheets/sheet1.xml", "xl/worksheets/_rels/sheet1.xml.rels");
            AssertNoDanglingRels(check);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void OpenSave_PreservesPivotCaches_WithRemappedRelId()
    {
        // P0-27: workbook.xml 丢 <pivotCaches> 会让 Excel 直接拒绝打开含透视表的文件
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets["Sheet1"].SetValue("A1", "v");
            wb.SaveAs(file);

            using (var zip = new ZipArchive(File.Open(file, FileMode.Open, FileAccess.ReadWrite), ZipArchiveMode.Update))
            {
                WriteEntry(zip, "xl/pivotCache/pivotCacheDefinition1.xml",
                    "<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" recordCount=\"0\"/>");
                AddWorkbookRel(zip, "rId900", $"{OfficeRelNs}/pivotCacheDefinition", "pivotCache/pivotCacheDefinition1.xml");
                AddContentTypeOverride(zip, "/xl/pivotCache/pivotCacheDefinition1.xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml");
                InsertIntoWorkbookXml(zip, "<calcPr fullCalcOnLoad=\"1\"/>",
                    "<pivotCaches><pivotCache cacheId=\"7\" r:id=\"rId900\"/></pivotCaches>");
            }

            var opened = Excel.Open(file);
            opened.Worksheets[0].SetValue("B1", "x");
            opened.Save();

            using var check = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            var savedWb = ReadText(check, "xl/workbook.xml");
            Assert.Matches("<pivotCaches[ >]", savedWb);
            Assert.Contains("cacheId=\"7\"", savedWb);

            // schema 全序：pivotCaches 必须在 calcPr 之后，否则 Excel 报 XML 错误
            int iCalc = savedWb.IndexOf("<calcPr", StringComparison.Ordinal);
            int iPivot = Regex.Match(savedWb, "<pivotCaches[ >]").Index;
            Assert.True(iCalc >= 0 && iPivot > iCalc, $"pivotCaches 位置错误: calcPr@{iCalc} pivotCaches@{iPivot}");

            // rId 重编号后必须指向真实的 pivotCacheDefinition
            var pivotRelId = Regex.Match(savedWb, "<pivotCache[^>]*r:id=\"([^\"]+)\"").Groups[1].Value;
            Assert.Contains("pivotCacheDefinition", RelTypeFor(check, "xl/_rels/workbook.xml.rels", pivotRelId));
            Assert.NotNull(check.GetEntry("xl/pivotCache/pivotCacheDefinition1.xml"));

            AssertRelIdsResolvable(check, "xl/workbook.xml", "xl/_rels/workbook.xml.rels");
            AssertNoDanglingRels(check);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void OpenSave_PreservesExternalReferences_WithRemappedRelId()
    {
        // P0-29: workbook.xml 丢 <externalReferences> 会让跨工作簿公式的缓存值与链接一起失效
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets["Sheet1"].SetValue("A1", "v");
            wb.SaveAs(file);

            using (var zip = new ZipArchive(File.Open(file, FileMode.Open, FileAccess.ReadWrite), ZipArchiveMode.Update))
            {
                WriteEntry(zip, "xl/externalLinks/externalLink1.xml",
                    "<externalLink xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>");
                AddWorkbookRel(zip, "rId901", $"{OfficeRelNs}/externalLink", "externalLinks/externalLink1.xml");
                AddContentTypeOverride(zip, "/xl/externalLinks/externalLink1.xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml");
                InsertIntoWorkbookXml(zip, "</sheets>",
                    "<externalReferences><externalReference r:id=\"rId901\"/></externalReferences>");
            }

            var opened = Excel.Open(file);
            opened.Worksheets[0].SetValue("B1", "x");
            opened.Save();

            using var check = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            var savedWb = ReadText(check, "xl/workbook.xml");
            Assert.Matches("<externalReferences[ >]", savedWb);

            // schema 全序：externalReferences 在 sheets 之后、definedNames / calcPr 之前
            int iSheets = savedWb.IndexOf("</sheets>", StringComparison.Ordinal);
            int iExt = Regex.Match(savedWb, "<externalReferences[ >]").Index;
            int iCalc = savedWb.IndexOf("<calcPr", StringComparison.Ordinal);
            Assert.True(iExt > iSheets, $"externalReferences 应在 sheets 之后: sheets@{iSheets} ext@{iExt}");
            Assert.True(iExt < iCalc, $"externalReferences 应在 calcPr 之前: ext@{iExt} calcPr@{iCalc}");

            var extRelId = Regex.Match(savedWb, "<externalReference[^>]*r:id=\"([^\"]+)\"").Groups[1].Value;
            Assert.Contains("externalLink", RelTypeFor(check, "xl/_rels/workbook.xml.rels", extRelId));
            Assert.NotNull(check.GetEntry("xl/externalLinks/externalLink1.xml"));

            AssertRelIdsResolvable(check, "xl/workbook.xml", "xl/_rels/workbook.xml.rels");
            AssertNoDanglingRels(check);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void OpenSave_DoesNotDuplicatePreservedImages()
    {
        // P0-28: 读取回填的图片已含在保真透传的 drawing 部件内，重复写出会让 open-save 图片翻倍
        var src = Path.Combine(AppContext.BaseDirectory, "Fixtures", "excel-authored-image-chart.xlsx");
        Assert.True(File.Exists(src), $"Required Excel image/chart fixture is missing: {src}");

        var file = GetTempFile();
        try
        {
            string drawingPath;
            int beforePic, beforeAnchor, beforeFrame, beforeMedia;
            using (var origin = new ZipArchive(File.OpenRead(src), ZipArchiveMode.Read))
            {
                drawingPath = origin.Entries
                    .Select(x => x.FullName)
                    .First(n => n.StartsWith("xl/drawings/drawing", StringComparison.Ordinal)
                             && n.EndsWith(".xml", StringComparison.Ordinal));
                var xml = ReadText(origin, drawingPath);
                beforePic = Regex.Matches(xml, "<xdr:pic\\b").Count;
                beforeAnchor = Regex.Matches(xml, "<xdr:twoCellAnchor\\b").Count;
                beforeFrame = Regex.Matches(xml, "<xdr:graphicFrame\\b").Count;
                beforeMedia = origin.Entries.Count(x => x.FullName.StartsWith("xl/media/", StringComparison.Ordinal));
            }
            Assert.True(beforePic > 0, "fixture 应含至少一张浮动图片");

            var opened = Excel.Open(src);
            opened.SaveAs(file);

            using var check = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            var saved = ReadText(check, drawingPath);
            Assert.Equal(beforePic, Regex.Matches(saved, "<xdr:pic\\b").Count);
            Assert.Equal(beforeAnchor, Regex.Matches(saved, "<xdr:twoCellAnchor\\b").Count);
            Assert.Equal(beforeFrame, Regex.Matches(saved, "<xdr:graphicFrame\\b").Count);
            Assert.Equal(beforeMedia, check.Entries.Count(x => x.FullName.StartsWith("xl/media/", StringComparison.Ordinal)));

            AssertNoDanglingRels(check);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void AddImage_ToPreservedDrawing_AppendsExactlyOne()
    {
        // 保留 drawing 的图片跳过写出后，新增图片仍须正常并入（不能连新图一起漏掉）
        var src = Path.Combine(AppContext.BaseDirectory, "Fixtures", "excel-authored-image-chart.xlsx");
        Assert.True(File.Exists(src), $"Required Excel image/chart fixture is missing: {src}");

        var file = GetTempFile();
        try
        {
            string drawingPath;
            int beforePic, beforeMedia;
            using (var origin = new ZipArchive(File.OpenRead(src), ZipArchiveMode.Read))
            {
                drawingPath = origin.Entries
                    .Select(x => x.FullName)
                    .First(n => n.StartsWith("xl/drawings/drawing", StringComparison.Ordinal)
                             && n.EndsWith(".xml", StringComparison.Ordinal));
                beforePic = Regex.Matches(ReadText(origin, drawingPath), "<xdr:pic\\b").Count;
                beforeMedia = origin.Entries.Count(x => x.FullName.StartsWith("xl/media/", StringComparison.Ordinal));
            }

            var opened = Excel.Open(src);
            opened.Worksheets[0].AddImage(MinimalPng, 10, 2);
            opened.SaveAs(file);

            using var check = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            var saved = ReadText(check, drawingPath);
            Assert.Equal(beforePic + 1, Regex.Matches(saved, "<xdr:pic\\b").Count);
            Assert.Equal(beforeMedia + 1, check.Entries.Count(x => x.FullName.StartsWith("xl/media/", StringComparison.Ordinal)));

            AssertNoDanglingRels(check);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Append_PreservesWorkbookLevelReferences()
    {
        // Append 路径同样要带上 workbook.xml 内的保留引用，否则追加一次就丢透视表缓存 / 命名区域
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets["Sheet1"].SetValue("A1", "v");
            wb.SaveAs(file);

            using (var zip = new ZipArchive(File.Open(file, FileMode.Open, FileAccess.ReadWrite), ZipArchiveMode.Update))
            {
                WriteEntry(zip, "xl/pivotCache/pivotCacheDefinition1.xml",
                    "<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" recordCount=\"0\"/>");
                AddWorkbookRel(zip, "rId900", $"{OfficeRelNs}/pivotCacheDefinition", "pivotCache/pivotCacheDefinition1.xml");
                AddContentTypeOverride(zip, "/xl/pivotCache/pivotCacheDefinition1.xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml");
                InsertIntoWorkbookXml(zip, "<calcPr fullCalcOnLoad=\"1\"/>",
                    "<pivotCaches><pivotCache cacheId=\"7\" r:id=\"rId900\"/></pivotCaches>");
                InsertIntoWorkbookXml(zip, "</sheets>",
                    "<definedNames><definedName name=\"区域\">Sheet1!$A$1</definedName></definedNames>");
            }

            var extra = new SheetData { SheetName = "Sheet1" };
            extra.Rows.Add(new List<Cell> { new Cell { Text = "追加" } });
            Excel.Append(file, extra);

            using var check = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            var savedWb = ReadText(check, "xl/workbook.xml");
            Assert.Matches("<pivotCaches[ >]", savedWb);
            Assert.Matches("<definedNames[ >]", savedWb);
            AssertRelIdsResolvable(check, "xl/workbook.xml", "xl/_rels/workbook.xml.rels");
            AssertNoDanglingRels(check);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    /// <summary>1x1 有效 PNG（Excel 可打开，手工构造的字节会被拒） </summary>
    private static byte[] MinimalPng => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mP8/x8AAwMB/6X5x2wAAAAASUVORK5CYII=");

    /// <summary>Q1: 工作表原始起始行号保真——数据从第 3 行起的文件 open-save 后必须仍从第 3 行起，
    /// 前导空行补齐，不能整体上移（透视表/合并区域等绝对行号引用会错位） </summary>
    [Fact]
    public void OpenSave_PreservesFirstRowNumber_WhenDataStartsAtOffset()
    {
        var file = GetTempFile();
        try
        {
            // 直接构造 sheet1.xml：数据从第 3 行起（前两行空占位），合并 B3:C3
            using (var zip = new ZipArchive(File.Open(file, FileMode.Create, FileAccess.ReadWrite), ZipArchiveMode.Update))
            {
                WriteEntry(zip, "xl/worksheets/sheet1.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                    "<sheetData>" +
                    "<row r=\"3\"><c r=\"A3\" t=\"s\"><v>0</v></c><c r=\"B3\"/><c r=\"C3\"/></row>" +
                    "<row r=\"4\"><c r=\"A4\"><v>1</v></c></row>" +
                    "<row r=\"5\"><c r=\"A5\"><v>2</v></c></row>" +
                    "</sheetData>" +
                    "<mergeCells><mergeCell ref=\"B3:C3\"/></mergeCells>" +
                    "</worksheet>");
                WriteEntry(zip, "xl/sharedStrings.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"3\" uniqueCount=\"1\">" +
                    "<si><t>头</t></si></sst>");
                WriteEntry(zip, "xl/styles.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><cellXfs count=\"1\"><xf numFmtId=\"0\"/></cellXfs></styleSheet>");
                WriteEntry(zip, "xl/workbook.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                    "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
                WriteEntry(zip, "xl/_rels/workbook.xml.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    $"<Relationships xmlns=\"{PkgRelNs}\">" +
                    $"<Relationship Id=\"rId1\" Type=\"{OfficeRelNs}/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
                    $"<Relationship Id=\"rId2\" Type=\"{OfficeRelNs}/sharedStrings\" Target=\"sharedStrings.xml\"/>" +
                    $"<Relationship Id=\"rId3\" Type=\"{OfficeRelNs}/styles\" Target=\"styles.xml\"/>" +
                    "</Relationships>");
                WriteEntry(zip, "_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    $"<Relationships xmlns=\"{PkgRelNs}\">" +
                    $"<Relationship Id=\"rId1\" Type=\"{OfficeRelNs}/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                    "</Relationships>");
                WriteEntry(zip, "[Content_Types].xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                    "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                    "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                    "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                    "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                    "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>" +
                    "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" +
                    "</Types>");
            }

            var opened = Excel.Open(file);
            Assert.Equal(3, opened.Worksheets[0].FirstRowNumber);
            // grid[0] 对应源行 3，B3 -> grid 0
            Assert.Equal(0, opened.Worksheets[0].MergedRanges[0].FirstRow);
            opened.Worksheets[0].SetValue("D1", "add");
            opened.Save();

            using var check = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            var saved = ReadText(check, "xl/worksheets/sheet1.xml");
            Assert.Contains("<row r=\"1\"", saved);
            Assert.Contains("<row r=\"2\"", saved);
            Assert.Contains("<c r=\"A3\"", saved);
            Assert.Contains("<c r=\"A5\"", saved);
            Assert.Matches("ref=\"B3:C3\"", saved);
            AssertNoDanglingRels(check);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    /// <summary>Q2: workbook extLst（含 slicerCaches）+ sheet extLst（含 slicerList）+ 原 sheetId
    /// 三者 open-save 后必须全部保留，r:id 经重编号后仍可解析。切片器缓存以 tabId 引用 sheetId，
    /// sheetId 被重排为 1-based 位置序号会导致切片器成孤儿（Excel COM SlicerCaches.Count=0）。</summary>
    [Fact]
    public void OpenSave_PreservesSlicerExtLstAndOriginalSheetId()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets["Sheet1"].SetValue("A1", "v");
            wb.SaveAs(file);

            using (var zip = new ZipArchive(File.Open(file, FileMode.Open, FileAccess.ReadWrite), ZipArchiveMode.Update))
            {
                // workbook.xml：sheetId 改成 7（非默认 1）；注入 workbook extLst 含 slicerCaches 引用 rId900/rId901
                var wbXml = ReadText(zip, "xl/workbook.xml");
                wbXml = Regex.Replace(wbXml, "sheetId=\"1\"", "sheetId=\"7\"");
                const string SlicerExtLst =
                    "<extLst><ext uri=\"{BBE1A952-AA13-448e-AADC-164F8A28A991}\" " +
                    "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">" +
                    "<x14:slicerCaches><x14:slicerCache r:id=\"rId900\"/>" +
                    "<x14:slicerCache r:id=\"rId901\"/></x14:slicerCaches></ext></extLst>";
                wbXml = wbXml.Replace("</workbook>", SlicerExtLst + "</workbook>");
                WriteEntry(zip, "xl/workbook.xml", wbXml);
                AddWorkbookRel(zip, "rId900", "http://schemas.microsoft.com/office/2007/relationships/slicerCache",
                    "slicerCaches/slicerCache1.xml");
                AddWorkbookRel(zip, "rId901", "http://schemas.microsoft.com/office/2007/relationships/slicerCache",
                    "slicerCaches/slicerCache2.xml");

                // slicer 部件本体（最小有效内容）
                WriteEntry(zip, "xl/slicerCaches/slicerCache1.xml",
                    "<slicerCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" " +
                    "name=\"X\" tabId=\"7\"><pivotTables/></slicerCacheDefinition>");
                WriteEntry(zip, "xl/slicerCaches/slicerCache2.xml",
                    "<slicerCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" " +
                    "name=\"Y\" tabId=\"7\"><pivotTables/></slicerCacheDefinition>");
                WriteEntry(zip, "xl/slicers/slicer1.xml",
                    "<slicer xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" name=\"X\" " +
                    "cache=\"slicerCache1\"><extLst/></slicer>");

                // sheet1.xml：注入 sheet extLst 含 slicerList 引用 rId800
                var sheet = ReadText(zip, "xl/worksheets/sheet1.xml");
                const string SheetExtLst =
                    "<extLst><ext uri=\"{A8765BA9-456A-4dab-B4F3-ACF838C121DE}\" " +
                    "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">" +
                    "<x14:slicerList><x14:slicer r:id=\"rId800\"/></x14:slicerList></ext></extLst>";
                sheet = sheet.Replace("</worksheet>", SheetExtLst + "</worksheet>");
                WriteEntry(zip, "xl/worksheets/sheet1.xml", sheet);

                // sheet1 rels：加 slicer rel rId800
                var srels = ReadText(zip, "xl/worksheets/_rels/sheet1.xml.rels");
                const string slicerRel =
                    $"<Relationship Id=\"rId800\" Type=\"http://schemas.microsoft.com/office/2007/relationships/slicer\" " +
                    "Target=\"../slicers/slicer1.xml\"/>";
                if (srels.Length == 0)
                    srels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                            $"<Relationships xmlns=\"{PkgRelNs}\">{slicerRel}</Relationships>";
                else if (srels.Contains("</Relationships>"))
                    srels = srels.Replace("</Relationships>", slicerRel + "</Relationships>");
                else
                    srels = srels.Replace("<Relationships/>",
                        $"<Relationships xmlns=\"{PkgRelNs}\">{slicerRel}</Relationships>");
                WriteEntry(zip, "xl/worksheets/_rels/sheet1.xml.rels", srels);

                AddContentTypeOverride(zip, "/xl/slicerCaches/slicerCache1.xml", "application/vnd.ms-excel.slicerCache+xml");
                AddContentTypeOverride(zip, "/xl/slicerCaches/slicerCache2.xml", "application/vnd.ms-excel.slicerCache+xml");
                AddContentTypeOverride(zip, "/xl/slicers/slicer1.xml", "application/vnd.ms-excel.slicer+xml");
            }

            var opened = Excel.Open(file);
            opened.Worksheets[0].SetValue("B1", "x");
            opened.Save();

            using var check = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            var savedWb = ReadText(check, "xl/workbook.xml");
            // sheetId 必须仍是 7（不能被重排为 1）
            Assert.Matches("sheetId=\"7\"", savedWb);
            // workbook extLst 含 2 个 slicerCache 引用，rId 经重编号后仍可解析
            Assert.Contains("slicerCaches", savedWb);
            AssertRelIdsResolvable(check, "xl/workbook.xml", "xl/_rels/workbook.xml.rels");

            var savedSheet = ReadText(check, "xl/worksheets/sheet1.xml");
            Assert.Contains("slicerList", savedSheet);
            AssertRelIdsResolvable(check, "xl/worksheets/sheet1.xml", "xl/worksheets/_rels/sheet1.xml.rels");

            // 部件本体仍存在
            Assert.NotNull(check.GetEntry("xl/slicerCaches/slicerCache1.xml"));
            Assert.NotNull(check.GetEntry("xl/slicerCaches/slicerCache2.xml"));
            Assert.NotNull(check.GetEntry("xl/slicers/slicer1.xml"));
            AssertNoDanglingRels(check);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
