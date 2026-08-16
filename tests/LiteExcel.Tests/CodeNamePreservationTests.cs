using LiteExcel;
using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace LiteExcel.Tests;

/// <summary>
/// 带宏工作簿的 codeName 绑定保留（见 issue #4）：
/// 打开时捕获 workbookPr@codeName / sheetPr@codeName，保存时写回，
/// 避免保留的 vbaProject 与宿主失去绑定后被 Excel 重排（ThisWorkbook1 / 事件宏失效）。
/// </summary>
public class CodeNamePreservationTests
{
    private static string GetTempFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"litexlsx_cn_{Guid.NewGuid():N}{ext}");

    private static string CreateBaseFile(string ext, int sheetCount)
    {
        var file = GetTempFile(ext);
        var wb = Excel.Create();
        for (int i = 1; i < sheetCount; i++)
            wb.Worksheets.Add($"表{i + 1}");
        for (int i = 0; i < sheetCount; i++)
            wb.Worksheets[i].SetValue("A1", $"v{i}");
        wb.SaveAs(file);
        return file;
    }

    /// <summary>向已生成文件注入 codeName：workbook.xml 加 workbookPr，指定 sheet 加 sheetPr（带无关属性与子元素，模拟真实文件） </summary>
    private static void InjectCodeNames(string file, string workbookCodeName, Dictionary<int, string> sheetCodeNames)
    {
        using var zip = new ZipArchive(File.Open(file, FileMode.Open, FileAccess.ReadWrite), ZipArchiveMode.Update);

        var wbEntry = zip.GetEntry("xl/workbook.xml")!;
        using (var s = wbEntry.Open())
        {
            var doc = XDocument.Load(s);
            var ns = doc.Root!.GetDefaultNamespace();
            doc.Root.AddFirst(new XElement(ns + "workbookPr", new XAttribute("codeName", workbookCodeName)));
            s.Position = 0;
            s.SetLength(0);
            doc.Save(s);
        }

        foreach (var (sheetIndex, codeName) in sheetCodeNames)
        {
            var entry = zip.GetEntry($"xl/worksheets/sheet{sheetIndex}.xml")!;
            using var s = entry.Open();
            var doc = XDocument.Load(s);
            var ns = doc.Root!.GetDefaultNamespace();
            // 模拟 Excel 真实文件：sheetPr 带无关属性 filterMode 与子元素 tabColor
            doc.Root.AddFirst(new XElement(ns + "sheetPr",
                new XAttribute("codeName", codeName),
                new XAttribute("filterMode", "false"),
                new XElement(ns + "tabColor", new XAttribute("rgb", "FFFF0000"))));
            s.Position = 0;
            s.SetLength(0);
            doc.Save(s);
        }
    }

    private static string? ReadSheetCodeName(string file, int sheetIndex)
    {
        using var zip = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
        var entry = zip.GetEntry($"xl/worksheets/sheet{sheetIndex}.xml");
        if (entry is null) return null;
        using var s = entry.Open();
        var doc = XDocument.Load(s);
        var ns = doc.Root!.GetDefaultNamespace();
        return doc.Root.Element(ns + "sheetPr")?.Attribute("codeName")?.Value;
    }

    private static string? ReadWorkbookCodeName(string file)
    {
        using var zip = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
        var entry = zip.GetEntry("xl/workbook.xml")!;
        using var s = entry.Open();
        var doc = XDocument.Load(s);
        var ns = doc.Root!.GetDefaultNamespace();
        return doc.Root.Element(ns + "workbookPr")?.Attribute("codeName")?.Value;
    }

    [Fact]
    public void Open_SaveAs_PreservesWorkbookAndSheetCodeNames()
    {
        var src = CreateBaseFile(".xlsx", 2);
        var dst = GetTempFile(".xlsx");
        try
        {
            InjectCodeNames(src, "ThisWorkbook", new Dictionary<int, string> { [1] = "Sheet1" });

            var opened = Excel.Open(src);
            opened.Worksheets[0].SetValue("B2", "改");
            opened.SaveAs(dst);

            Assert.Equal("ThisWorkbook", ReadWorkbookCodeName(dst));
            Assert.Equal("Sheet1", ReadSheetCodeName(dst, 1));
        }
        finally { if (File.Exists(src)) File.Delete(src); if (File.Exists(dst)) File.Delete(dst); }
    }

    [Fact]
    public void NoCodeNames_Source_OutputStaysClean()
    {
        var src = CreateBaseFile(".xlsx", 1);
        var dst = GetTempFile(".xlsx");
        try
        {
            var opened = Excel.Open(src);
            opened.SaveAs(dst);

            Assert.Null(ReadWorkbookCodeName(dst));
            Assert.Null(ReadSheetCodeName(dst, 1));
        }
        finally { if (File.Exists(src)) File.Delete(src); if (File.Exists(dst)) File.Delete(dst); }
    }

    [Fact]
    public void MultiSheet_CodeNamesStayWithTheirSheets()
    {
        var src = CreateBaseFile(".xlsx", 2);
        var dst = GetTempFile(".xlsx");
        try
        {
            InjectCodeNames(src, "ThisWorkbook", new Dictionary<int, string> { [1] = "SheetA", [2] = "SheetZ" });

            var opened = Excel.Open(src);
            opened.SaveAs(dst);

            Assert.Equal("SheetA", ReadSheetCodeName(dst, 1));
            Assert.Equal("SheetZ", ReadSheetCodeName(dst, 2));
        }
        finally { if (File.Exists(src)) File.Delete(src); if (File.Exists(dst)) File.Delete(dst); }
    }

    [Fact]
    public void AddedSheet_HasNoSheetPr_ExistingSheetsKeepTheirs()
    {
        var src = CreateBaseFile(".xlsx", 2);
        var dst = GetTempFile(".xlsx");
        try
        {
            InjectCodeNames(src, "ThisWorkbook", new Dictionary<int, string> { [1] = "SheetA" });

            var opened = Excel.Open(src);
            opened.Worksheets.Add("新表");
            opened.SaveAs(dst);

            Assert.Equal("SheetA", ReadSheetCodeName(dst, 1));
            Assert.Equal("ThisWorkbook", ReadWorkbookCodeName(dst));
            // 新表从未有 codeName，不应凭空产生 sheetPr
            using var zip = new ZipArchive(File.OpenRead(dst), ZipArchiveMode.Read);
            var entry = zip.GetEntry("xl/worksheets/sheet3.xml")!;
            using var s = entry.Open();
            var doc = XDocument.Load(s);
            var ns = doc.Root!.GetDefaultNamespace();
            Assert.Null(doc.Root.Element(ns + "sheetPr"));
        }
        finally { if (File.Exists(src)) File.Delete(src); if (File.Exists(dst)) File.Delete(dst); }
    }

    [Fact]
    public void SheetPr_WithExtraContent_OnlyCodeNameWrittenBack()
    {
        var src = CreateBaseFile(".xlsx", 1);
        var dst = GetTempFile(".xlsx");
        try
        {
            InjectCodeNames(src, "ThisWorkbook", new Dictionary<int, string> { [1] = "Sheet1" });

            var opened = Excel.Open(src);
            opened.SaveAs(dst);

            using var zip = new ZipArchive(File.OpenRead(dst), ZipArchiveMode.Read);
            var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!;
            using var s = entry.Open();
            var doc = XDocument.Load(s);
            var ns = doc.Root!.GetDefaultNamespace();
            var sheetPr = doc.Root.Element(ns + "sheetPr");
            Assert.NotNull(sheetPr);
            Assert.Equal("Sheet1", sheetPr!.Attribute("codeName")?.Value);
            // 无关属性与子元素不透传（当前只保留绑定所需的 codeName）
            Assert.Null(sheetPr.Attribute("filterMode"));
            Assert.Empty(sheetPr.Elements());
        }
        finally { if (File.Exists(src)) File.Delete(src); if (File.Exists(dst)) File.Delete(dst); }
    }
}
