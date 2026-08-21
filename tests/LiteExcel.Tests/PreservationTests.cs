using LiteExcel;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace LiteExcel.Tests;

/// <summary>
/// 打开-修改-保存 时未映射 OOXML 部件（宏/主题/绘图/自定义 rels）的保留行为。
/// </summary>
public class PreservationTests
{
    private static readonly byte[] FakeVba = { 0xCC, 0xFE, 0xED, 0x01, 0x02, 0x03, 0x00, 0xFF };

    private static string GetTempFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"litexlsx_pres_{Guid.NewGuid():N}{ext}");

    /// <summary>向已生成的 xlsx/xlsm 注入 vbaProject.bin、theme、drawing 部件及其 rels / content types 声明 </summary>
    private static void InjectExtraParts(string file)
    {
        using var zip = new ZipArchive(File.Open(file, FileMode.Open, FileAccess.ReadWrite), ZipArchiveMode.Update);

        var vba = zip.CreateEntry("xl/vbaProject.bin");
        using (var s = vba.Open()) s.Write(FakeVba, 0, FakeVba.Length);

        var theme = zip.CreateEntry("xl/theme/theme1.xml");
        using (var s = theme.Open())
        {
            var bytes = Encoding.UTF8.GetBytes("<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\"/>");
            s.Write(bytes, 0, bytes.Length);
        }

        var drawing = zip.CreateEntry("xl/drawings/drawing1.xml");
        using (var s = drawing.Open())
        {
            var bytes = Encoding.UTF8.GetBytes("<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>");
            s.Write(bytes, 0, bytes.Length);
        }

        var wbRels = zip.GetEntry("xl/_rels/workbook.xml.rels")!;
        using (var s = wbRels.Open())
        {
            var doc = XDocument.Load(s);
            var ns = doc.Root!.GetDefaultNamespace();
            doc.Root.Add(new XElement(ns + "Relationship",
                new XAttribute("Id", "rId901"),
                new XAttribute("Type", "http://schemas.microsoft.com/office/2006/relationships/vbaProject"),
                new XAttribute("Target", "vbaProject.bin")));
            doc.Root.Add(new XElement(ns + "Relationship",
                new XAttribute("Id", "rId902"),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"),
                new XAttribute("Target", "theme/theme1.xml")));
            s.Position = 0;
            s.SetLength(0);
            doc.Save(s);
        }

        var sheetRels = zip.CreateEntry("xl/worksheets/_rels/sheet1.xml.rels");
        using (var s = sheetRels.Open())
        {
            var rels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing1.xml\"/>" +
                "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com\" TargetMode=\"External\"/>" +
                "</Relationships>";
            var bytes = Encoding.UTF8.GetBytes(rels);
            s.Write(bytes, 0, bytes.Length);
        }

        var ct = zip.GetEntry("[Content_Types].xml")!;
        using (var s = ct.Open())
        {
            var doc = XDocument.Load(s);
            var ns = doc.Root!.GetDefaultNamespace();
            doc.Root.Add(new XElement(ns + "Default",
                new XAttribute("Extension", "bin"),
                new XAttribute("ContentType", "application/vnd.ms-office.vbaProject")));
            doc.Root.Add(new XElement(ns + "Override",
                new XAttribute("PartName", "/xl/theme/theme1.xml"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.theme+xml")));
            doc.Root.Add(new XElement(ns + "Override",
                new XAttribute("PartName", "/xl/drawings/drawing1.xml"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml")));
            s.Position = 0;
            s.SetLength(0);
            doc.Save(s);
        }
    }

    private static string ReadText(ZipArchiveEntry entry)
    {
        using var s = entry.Open();
        using var r = new StreamReader(s, Encoding.UTF8);
        return r.ReadToEnd();
    }

    private static byte[] ReadBytes(ZipArchiveEntry entry)
    {
        using var s = entry.Open();
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        return ms.ToArray();
    }

    [Fact]
    public void Open_Modify_Save_PreservesUnknownParts_And_Data()
    {
        // 注入 vbaProject.bin 的文件必须保存为 xlsm（xlsx 不支持宏）
        var file = GetTempFile(".xlsm");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsm);
            wb.Worksheets["Sheet1"].SetValue("A1", "数据");
            wb.SaveAs(file);
            InjectExtraParts(file);

            var opened = Excel.Open(file);
            opened.Worksheets[0].SetValue("B1", "新增");
            opened.Save();

            // 值往返
            var reopened = Excel.Open(file);
            Assert.Equal("数据", reopened.Worksheets[0].Cell("A1").GetString());
            Assert.Equal("新增", reopened.Worksheets[0].Cell("B1").GetString());

            using var zip = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            // 未映射部件按 blob 保留
            var vba = zip.GetEntry("xl/vbaProject.bin");
            Assert.NotNull(vba);
            Assert.Equal(FakeVba, ReadBytes(vba!));
            Assert.NotNull(zip.GetEntry("xl/theme/theme1.xml"));
            Assert.NotNull(zip.GetEntry("xl/drawings/drawing1.xml"));

            // 工作簿 rels 合并：vbaProject + theme 关系保留
            var wbRels = ReadText(zip.GetEntry("xl/_rels/workbook.xml.rels")!);
            Assert.Contains("/vbaProject", wbRels);
            Assert.Contains("theme/theme1.xml", wbRels);

            // 工作表 rels 合并：drawing + 外部超链接保留
            var sheetRels = ReadText(zip.GetEntry("xl/worksheets/_rels/sheet1.xml.rels")!);
            Assert.Contains("drawings/drawing1.xml", sheetRels);
            Assert.Contains("TargetMode=\"External\"", sheetRels);

            // content types 保留 bin / theme / drawing 声明
            var ct = ReadText(zip.GetEntry("[Content_Types].xml")!);
            Assert.Contains("vbaProject", ct);
            Assert.Contains("theme1.xml", ct);
            Assert.Contains("drawing1.xml", ct);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Open_Modify_Save_PreservesMacroPart_Xlsm()
    {
        var file = GetTempFile(".xlsm");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsm);
            wb.Worksheets["Sheet1"].SetValue("A1", "宏表");
            wb.SaveAs(file);
            InjectExtraParts(file);

            var opened = Excel.Open(file);
            Assert.Equal(ExcelFormat.Xlsm, opened.Format);
            opened.Worksheets[0].SetValue("A2", "x");
            opened.Save();

            using var zip = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            var vba = zip.GetEntry("xl/vbaProject.bin");
            Assert.NotNull(vba);
            Assert.Equal(FakeVba, ReadBytes(vba!));

            var wbRels = ReadText(zip.GetEntry("xl/_rels/workbook.xml.rels")!);
            Assert.Contains("/vbaProject", wbRels);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void RenameSheet_PreservesDrawingRelation()
    {
        // P0-3: 改表名不应丢弃 drawing/图片关联（表名与 sheet rels 无绑定关系）
        var file = GetTempFile(".xlsm");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsm);
            wb.Worksheets["Sheet1"].SetValue("A1", "v");
            wb.SaveAs(file);
            InjectExtraParts(file);

            var opened = Excel.Open(file);
            opened.Worksheets[0].Name = "改名"; // 仅改名，结构数量不变
            opened.Save();

            using var zip = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            // drawing 部件保留
            Assert.NotNull(zip.GetEntry("xl/drawings/drawing1.xml"));
            // 工作表级保留 rels 合并：drawing 关联仍在
            var sheetRels = zip.GetEntry("xl/worksheets/_rels/sheet1.xml.rels");
            Assert.NotNull(sheetRels);
            var text = ReadText(sheetRels!);
            Assert.Contains("drawings/drawing1.xml", text);
            // 宏仍保留
            Assert.NotNull(zip.GetEntry("xl/vbaProject.bin"));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void AddRemoveSheet_DropsSheetRels_ButKeepsBlobs()
    {
        // 数量变化（增/删 sheet）时不再合并工作表级保留 rels，但 blob 部件仍保留（无害孤儿）
        var file = GetTempFile(".xlsm");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsm);
            wb.Worksheets["Sheet1"].SetValue("A1", "v");
            wb.SaveAs(file);
            InjectExtraParts(file);

            var opened = Excel.Open(file);
            opened.Worksheets.Add("NewSheet");
            opened.Save();

            using var zip = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            // 数量变化 → 丢弃 sheet1 的保留 drawing rel，但 drawing blob 仍保留
            var sheetRels = zip.GetEntry("xl/worksheets/_rels/sheet1.xml.rels");
            if (sheetRels is not null)
            {
                var text = ReadText(sheetRels);
                Assert.DoesNotContain("drawings/drawing1.xml", text);
            }
            Assert.NotNull(zip.GetEntry("xl/drawings/drawing1.xml"));
            Assert.NotNull(zip.GetEntry("xl/vbaProject.bin"));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void NewWorkbook_Save_HasNoPreservedParts()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            wb.Worksheets["Sheet1"].SetValue("A1", "1");
            wb.SaveAs(file);

            using var zip = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            Assert.Null(zip.GetEntry("xl/vbaProject.bin"));
            Assert.Null(zip.GetEntry("xl/worksheets/_rels/sheet1.xml.rels"));

            // rels 仅含写入器固有条目
            var wbRels = ReadText(zip.GetEntry("xl/_rels/workbook.xml.rels")!);
            Assert.DoesNotContain("vbaProject", wbRels);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Append_PreservesMacroAndChartParts()
    {
        // P0-25: Append 须透传保留部件，xlsm 追加数据不丢宏
        var file = GetTempFile(".xlsm");
        try
        {
            var orig = new SheetData
            {
                SheetName = "Sheet1",
                Headers = new() { "Col" },
                Rows = new() { new[] { Cell.FromText("v") } },
            };
            XlsxWriter.Write(file, orig);
            InjectExtraParts(file);

            var appendData = new SheetData
            {
                SheetName = "Sheet1",
                Headers = new() { "Col" },
                Rows = new() { new[] { Cell.FromText("追加") } },
            };
            XlsxWriter.Append(file, appendData);

            using var zip = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            // 宏部件与 theme/drawing 保留
            var vba = zip.GetEntry("xl/vbaProject.bin");
            Assert.NotNull(vba);
            Assert.Equal(FakeVba, ReadBytes(vba!));
            Assert.NotNull(zip.GetEntry("xl/theme/theme1.xml"));
            Assert.NotNull(zip.GetEntry("xl/drawings/drawing1.xml"));

            // workbook rels 合并：vbaProject + theme 关系保留
            var wbRels = ReadText(zip.GetEntry("xl/_rels/workbook.xml.rels")!);
            Assert.Contains("/vbaProject", wbRels);
            Assert.Contains("theme/theme1.xml", wbRels);

            // 追加的数据在（表头行 1，原数据行 2，追加行 3）
            var reopened = Excel.Open(file);
            Assert.Equal("追加", reopened.Worksheets[0].Cell("A3").GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void InCell_OpenThenAddImage_NoDuplicateZipEntry()
    {
        // P0-11: 打开含 InCell richData 的文件再加图，保存时保留部件不与其重建条目重名
        var file = GetTempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].Cell("A1").SetValue("InCell");
            wb.Worksheets[0].AddImage(TestPng, 2, 1, placement: ImagePlacement.InCell);
            wb.SaveAs(file);

            var opened = Excel.Open(file);
            opened.Worksheets[0].AddImage(TestPng, 5, 1, placement: ImagePlacement.InCell);
            opened.Save();

            using var zip = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            Assert.NotNull(zip.GetEntry("xl/metadata.xml"));
            Assert.NotNull(zip.GetEntry("xl/richData/richValueRel.xml"));
            Assert.NotNull(zip.GetEntry("xl/richData/rdrichvalue.xml"));
            Assert.NotNull(zip.GetEntry("xl/richData/rdrichvaluestructure.xml"));
            Assert.NotNull(zip.GetEntry("xl/richData/rdRichValueTypes.xml"));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void RealExcelFixture_RoundTrip_PreservesTableThemeCustom()
    {
        // P0-5: 用真实 Excel 样本（含表格/主题/自定义属性）做打开-修改-保存往返
        var fixture = Path.Combine(AppContext.BaseDirectory, "Fixtures", "excel-authored-compatibility.xlsx");
        Assert.True(File.Exists(fixture), $"缺少真实样本: {fixture}");

        var file = GetTempFile(".xlsx");
        try
        {
            File.Copy(fixture, file, overwrite: true);
            var opened = Excel.Open(file);
            Assert.True(opened.Worksheets.Count >= 2);
            opened.Worksheets[0].SetValue("B2", "改动");
            opened.Save();

            using var zip = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            // 真实部件保留：表格/主题/自定义属性
            Assert.NotNull(zip.GetEntry("xl/tables/table1.xml"));
            Assert.NotNull(zip.GetEntry("xl/theme/theme1.xml"));
            Assert.NotNull(zip.GetEntry("docProps/custom.xml"));
            // 陈旧 calcChain 不透传（P0-12 联动）
            Assert.Null(zip.GetEntry("xl/calcChain.xml"));

            // 数据仍可读
            var reopened = Excel.Open(file);
            Assert.Equal("改动", reopened.Worksheets[0].Cell("B2").GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    private static readonly byte[] TestPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");
}
