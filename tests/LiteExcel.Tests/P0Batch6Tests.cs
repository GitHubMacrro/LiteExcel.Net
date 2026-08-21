using LiteExcel;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace LiteExcel.Tests;

/// <summary>
/// 批次 6：流式写入器补齐 P0-20/21/22/23/24。
/// </summary>
public class P0Batch6Tests
{
    private static string GetTempFile(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"p0b6_{Guid.NewGuid():N}{ext}");
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        return path;
    }

    private static string ReadZipEntry(string file, string entry)
    {
        using var zip = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
        using var s = zip.GetEntry(entry)!.Open();
        using var r = new StreamReader(s, Encoding.UTF8);
        return r.ReadToEnd();
    }

    // ── P0-23: 非法扩展名显式报错 ──

    [Theory]
    [InlineData(".csv")]
    [InlineData(".xls")]
    [InlineData(".xlsb")]
    public void P0_23_CreateWriter_RejectsNonXlsxExtensions(string ext)
    {
        var file = GetTempFile(ext);
        try
        {
            var ex = Assert.Throws<LiteExcelException>(() => Excel.CreateWriter(file));
            Assert.Contains("不支持流式写入", ex.Message);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void P0_23_StreamRows_RejectsNonXlsxFormats()
    {
        var file = GetTempFile(".csv");
        File.WriteAllText(file, "a,b\n1,2\n");
        try
        {
            var ex = Assert.Throws<LiteExcelException>(() =>
                Excel.StreamRows(file, "Sheet1", _ => { }));
            Assert.Contains("不支持流式读取", ex.Message);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void P0_23_Append_RejectsNonXlsxFormats()
    {
        var file = GetTempFile(".csv");
        File.WriteAllText(file, "a,b\n1,2\n");
        try
        {
            var sd = new SheetData { SheetName = "S1", Headers = new() { "a" }, Rows = new() { new[] { Cell.FromText("x") } } };
            var ex = Assert.Throws<LiteExcelException>(() => XlsxWriter.Append(file, sd));
            Assert.Contains("不支持追加", ex.Message);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void P0_23_CreateWriter_AcceptsXlsm()
    {
        var file = GetTempFile(".xlsm");
        try
        {
            using var w = Excel.CreateWriter(file);
            w.WriteRow(new object?[] { 1 });
            w.Close();
            // 验证 ContentTypes 写 macroEnabled
            var ct = ReadZipEntry(file, "[Content_Types].xml");
            Assert.Contains("macroEnabled", ct);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    // ── P0-20: 样式随行写出 ──

    [Fact]
    public void P0_20_Styles_Written_Back()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            using (var w = Excel.CreateWriter(file))
            {
                var boldStyle = new CellStyle { Bold = true };
                w.WriteRow(new Cell[]
                {
                    Cell.FromText("Header").WithStyle(boldStyle),
                    Cell.FromNumber(42, "0.00"),
                });
                w.WriteRow(new Cell[]
                {
                    Cell.FromText("x"),
                    Cell.FromNumber(3.14, "0.000"),
                });
                w.Close();
            }

            // styles.xml 应含多个 cellXfs（非仅默认一个）
            var stylesXml = ReadZipEntry(file, "xl/styles.xml");
            Assert.Contains("cellXfs", stylesXml);
            // bold 的 font 边应当有 <b/>
            Assert.Contains("<b/>", stylesXml);
            // 应有 numFmt（0.00 / 0.000）
            Assert.Contains("numFmt", stylesXml);

            // 单元格应带 s 属性
            var sheetXml = ReadZipEntry(file, "xl/worksheets/sheet1.xml");
            Assert.Contains(" s=\"", sheetXml);

            // 库能读回
            var wb = Excel.Open(file);
            Assert.Equal("Header", wb.Worksheets[0].Cell("A1").GetString());
            Assert.Equal(42.0, wb.Worksheets[0].Cell("B1").GetDouble());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    // ── P0-21: 公式写出 ──

    [Fact]
    public void P0_21_Formula_Written()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            using (var w = Excel.CreateWriter(file))
            {
                w.WriteRow(new Cell[]
                {
                    Cell.FromNumber(1),
                    Cell.FromNumber(2),
                    Cell.FromFormula("A1+B1"),
                });
                w.Close();
            }

            var sheetXml = ReadZipEntry(file, "xl/worksheets/sheet1.xml");
            Assert.Contains("<f>", sheetXml);
            Assert.Contains("A1+B1", sheetXml);

            // 库读回：公式保留
            var wb = Excel.Open(file);
            var c = wb.Worksheets[0].Cell("C1");
            Assert.True(c.IsFormula || !string.IsNullOrEmpty(c.Formula));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    // ── P0-22: 超链接写出 ──

    [Fact]
    public void P0_22_Hyperlinks_Written()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            using (var w = Excel.CreateWriter(file))
            {
                w.WriteRow(new Cell[]
                {
                    Cell.FromText("link").WithHyperlink(new Hyperlink { Target = "https://example.com", Tooltip = "tip" }),
                });
                w.WriteRow(new Cell[]
                {
                    Cell.FromText("internal").WithHyperlink(new Hyperlink { Target = "#Sheet1!A1", IsInternal = true }),
                });
                w.Close();
            }

            var sheetXml = ReadZipEntry(file, "xl/worksheets/sheet1.xml");
            Assert.Contains("<hyperlinks>", sheetXml);
            Assert.Contains("rIdH1", sheetXml);
            Assert.Contains("location=", sheetXml);

            // sheet rels 应有外部链接
            using var zip = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            var relsEntry = zip.GetEntry("xl/worksheets/_rels/sheet1.xml.rels");
            Assert.NotNull(relsEntry);
            using var rs = relsEntry!.Open();
            using var rr = new StreamReader(rs, Encoding.UTF8);
            var relsXml = rr.ReadToEnd();
            Assert.Contains("hyperlink", relsXml);
            Assert.Contains("example.com", relsXml);

            // 库读回
            var wb = Excel.Open(file);
            var h = wb.Worksheets[0].Cell("A1").Hyperlink;
            Assert.NotNull(h);
            Assert.Equal("https://example.com", h!.Target);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}

internal static class TestCellExtensions
{
    public static Cell WithStyle(this Cell c, CellStyle s) { c.Style = s; return c; }
    public static Cell WithHyperlink(this Cell c, Hyperlink h) { c.Hyperlink = h; return c; }
}
