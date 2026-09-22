using LiteExcel;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace LiteExcel.Tests;

/// <summary>
/// 回归：手术式删表（xlsb→xlsb）后包结构必须完整——不得残留「孤儿 .rels」
/// （.rels 的父部件已被删除，但 .rels 自身未删）或悬空引用。
/// 旧实现删被删表的依赖部件（drawingN.xml 等）本体时未一并删其 .rels，
/// 导致删含图片/图表 drawing 的表后 Excel 打开报修复或直接闪退。
/// </summary>
public class SurgicalDeleteIntegrityTests
{
    private const string PkgRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string OfficeRelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string DrawingCt = "application/vnd.openxmlformats-officedocument.drawing+xml";

    private static string GetTempFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"surgdel_{Guid.NewGuid():N}{ext}");

    /// <summary>构造：3 张表，其中 S1 带 drawing + 其 .rels → media；S0/S2 纯数据。</summary>
    private static string BuildWithDrawingSheet()
    {
        var spec = new XlsbTestFile.WorkbookSpec();
        spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "S0" });
        spec.Sheets[0].Rows.Add(new XlsbTestFile.RowSpec { Cells = { new XlsbTestFile.CellSpec { Col = 0, Text = "keep0" } } });
        spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "S1" });
        spec.Sheets[1].Rows.Add(new XlsbTestFile.RowSpec { Cells = { new XlsbTestFile.CellSpec { Col = 0, Text = "drop" } } });
        spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "S2" });
        spec.Sheets[2].Rows.Add(new XlsbTestFile.RowSpec { Cells = { new XlsbTestFile.CellSpec { Col = 0, Text = "keep2" } } });

        // S1 的 drawing 部件 + 其 .rels（指向 media/image1.png）
        spec.ExtraParts["xl/drawings/drawing1.xml"] =
            Encoding.UTF8.GetBytes("<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>");
        spec.ExtraParts["xl/media/image1.png"] = new byte[] { 0x89, 0x50, 0x4E, 0x47 };
        spec.ExtraOverrides["/xl/drawings/drawing1.xml"] = DrawingCt;
        spec.ExtraDefaults["png"] = "image/png";
        spec.ExtraRels["xl/worksheets/_rels/sheet2.bin.rels"] =
            $"<Relationship Id=\"rIdD\" Type=\"{OfficeRelNs}/drawing\" Target=\"../drawings/drawing1.xml\"/>";
        spec.ExtraRels["xl/drawings/_rels/drawing1.xml.rels"] =
            $"<Relationship Id=\"rId1\" Type=\"{OfficeRelNs}/image\" Target=\"../media/image1.png\"/>";
        return XlsbTestFile.Build(spec);
    }

    private static string ReadText(ZipArchive zip, string entry)
    {
        var e = zip.GetEntry(entry);
        if (e is null) return "";
        using var s = e.Open();
        using var r = new StreamReader(s, Encoding.UTF8);
        return r.ReadToEnd();
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

    private static List<string> FindOrphanRels(ZipArchive zip)
    {
        var names = new HashSet<string>(zip.Entries.Select(x => x.FullName), StringComparer.Ordinal);
        var orphans = new List<string>();
        foreach (var e in zip.Entries.Where(x => x.FullName.EndsWith(".rels", StringComparison.Ordinal)))
        {
            if (e.FullName == "_rels/.rels") continue;
            int idx = e.FullName.LastIndexOf("_rels/", StringComparison.Ordinal);
            string dir = idx > 0 ? e.FullName.Substring(0, idx) : "";
            string file = e.FullName.Substring(idx + "_rels/".Length);
            file = file.Substring(0, file.Length - ".rels".Length);
            if (!names.Contains(dir + file)) orphans.Add(e.FullName);
        }
        return orphans;
    }

    private static List<string> FindDanglingRels(ZipArchive zip)
    {
        var names = new HashSet<string>(zip.Entries.Select(x => x.FullName), StringComparer.Ordinal);
        var dangling = new List<string>();
        foreach (var e in zip.Entries.Where(x => x.FullName.EndsWith(".rels", StringComparison.Ordinal)))
        {
            var doc = XDocument.Parse(ReadText(zip, e.FullName));
            var ns = doc.Root!.GetDefaultNamespace();
            foreach (var rel in doc.Root.Elements(ns + "Relationship"))
            {
                if ((string?)rel.Attribute("TargetMode") == "External") continue;
                var target = (string?)rel.Attribute("Target") ?? "";
                var abs = ResolveRelTarget(e.FullName, target);
                if (!names.Contains(abs)) dangling.Add($"{e.FullName} -> {abs}");
            }
        }
        return dangling;
    }

    [Fact]
    public void DeleteSheetWithDrawing_LeavesNoOrphanRels()
    {
        var file = BuildWithDrawingSheet();
        var outPath = GetTempFile(".xlsb");
        try
        {
            var wb = Excel.Open(file);
            Assert.Equal(3, wb.Worksheets.Count);
            wb.Worksheets.First(w => w.Name == "S1").Delete();
            Assert.Equal(2, wb.Worksheets.Count);
            wb.SaveAs(outPath, ExcelFormat.Xlsb);

            using var zip = ZipFile.OpenRead(outPath);

            // 被删表依赖的 drawing 及其 .rels 都应被清除
            Assert.Null(zip.GetEntry("xl/drawings/drawing1.xml"));
            Assert.Null(zip.GetEntry("xl/drawings/_rels/drawing1.xml.rels"));

            var orphans = FindOrphanRels(zip);
            Assert.True(orphans.Count == 0, "存在孤儿 rels: " + string.Join(" | ", orphans));

            var dangling = FindDanglingRels(zip);
            Assert.True(dangling.Count == 0, "存在悬空 rel: " + string.Join(" | ", dangling));

            // 保留表数据仍在
            var reopened = Excel.Open(outPath);
            Assert.Equal(2, reopened.Worksheets.Count);
            Assert.Equal("S0", reopened.Worksheets[0].Name);
            Assert.Equal("S2", reopened.Worksheets[1].Name);
        }
        finally
        {
            if (File.Exists(file)) File.Delete(file);
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    [Fact]
    public void DeleteMultipleSheetsWithDrawings_LeavesNoOrphanRels()
    {
        // 4 张表，S1/S2 各带 drawing + .rels；连续删 S1、S2 → 不得残留任何孤儿 rels。
        var spec = new XlsbTestFile.WorkbookSpec();
        for (int i = 0; i < 4; i++)
            spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "S" + i });
        for (int i = 1; i <= 2; i++)
        {
            spec.ExtraParts[$"xl/drawings/drawing{i}.xml"] =
                Encoding.UTF8.GetBytes("<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>");
            spec.ExtraParts[$"xl/media/image{i}.png"] = new byte[] { 0x89, 0x50, 0x4E, 0x47 };
            spec.ExtraOverrides[$"/xl/drawings/drawing{i}.xml"] = DrawingCt;
            spec.ExtraRels[$"xl/worksheets/_rels/sheet{i + 1}.bin.rels"] =
                $"<Relationship Id=\"rIdD\" Type=\"{OfficeRelNs}/drawing\" Target=\"../drawings/drawing{i}.xml\"/>";
            spec.ExtraRels[$"xl/drawings/_rels/drawing{i}.xml.rels"] =
                $"<Relationship Id=\"rId1\" Type=\"{OfficeRelNs}/image\" Target=\"../media/image{i}.png\"/>";
        }
        spec.ExtraDefaults["png"] = "image/png";

        var file = XlsbTestFile.Build(spec);
        var outPath = GetTempFile(".xlsb");
        try
        {
            var wb = Excel.Open(file);
            Assert.Equal(4, wb.Worksheets.Count);
            wb.Worksheets.First(w => w.Name == "S1").Delete();
            wb.Worksheets.First(w => w.Name == "S2").Delete();
            Assert.Equal(2, wb.Worksheets.Count);
            wb.SaveAs(outPath, ExcelFormat.Xlsb);

            using var zip = ZipFile.OpenRead(outPath);
            var orphans = FindOrphanRels(zip);
            Assert.True(orphans.Count == 0, "存在孤儿 rels: " + string.Join(" | ", orphans));
            var dangling = FindDanglingRels(zip);
            Assert.True(dangling.Count == 0, "存在悬空 rel: " + string.Join(" | ", dangling));
        }
        finally
        {
            if (File.Exists(file)) File.Delete(file);
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }
}
