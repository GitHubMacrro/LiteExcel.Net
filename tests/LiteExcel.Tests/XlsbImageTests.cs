using LiteExcel;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace LiteExcel.Tests;

/// <summary>
/// xlsb 浮动图片写入：drawing/media 与 xlsx 同为 XML 部件；工作表经 BrtDrawing 引用 sheet rels 的 drawing 关系。
/// </summary>
public class XlsbImageTests
{
    private static readonly byte[] Png1x1 = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");

    private static string GetTempFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"xlsbimg_{Guid.NewGuid():N}{ext}");

    private static string ReadEntry(string file, string entry)
    {
        using var zip = ZipFile.OpenRead(file);
        var e = zip.GetEntry(entry);
        if (e is null) return "";
        using var r = new StreamReader(e.Open());
        return r.ReadToEnd();
    }

    private static byte[] ReadEntryBytes(string file, string entry)
    {
        using var zip = ZipFile.OpenRead(file);
        var e = zip.GetEntry(entry);
        if (e is null) return Array.Empty<byte>();
        using var ms = new MemoryStream();
        using (var s = e.Open()) s.CopyTo(ms);
        return ms.ToArray();
    }

    private static bool HasEntry(string file, string entry)
    {
        using var zip = ZipFile.OpenRead(file);
        return zip.GetEntry(entry) is not null;
    }

    [Fact]
    public void Xlsb_FloatingImage_WritesMediaDrawingAndBrtDrawing()
    {
        var file = GetTempFile(".xlsb");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "图片");
            ws.AddImage(Png1x1, 2, 1, widthPx: 50, heightPx: 50, placement: ImagePlacement.Floating);
            wb.SaveAs(file);

            Assert.Equal(Png1x1, ReadEntryBytes(file, "xl/media/image1.png"));
            var drawing = ReadEntry(file, "xl/drawings/drawing1.xml");
            Assert.Contains("oneCellAnchor", drawing);
            Assert.Contains("r:embed=\"rId1\"", drawing);
            Assert.True(HasEntry(file, "xl/drawings/_rels/drawing1.xml.rels"));

            // sheet rels 含 drawing 关系
            var sheetRels = ReadEntry(file, "xl/worksheets/_rels/sheet1.bin.rels");
            Assert.Contains("/relationships/drawing", sheetRels);
            Assert.Contains("../drawings/drawing1.xml", sheetRels);

            // sheet1.bin 含 BrtDrawing (0x0226)
            var sheetBin = ReadEntryBytes(file, "xl/worksheets/sheet1.bin");
            Assert.Contains(0x226, RecordTypes(sheetBin));

            // CT 声明 media Default + drawing Override
            var ct = ReadEntry(file, "[Content_Types].xml");
            Assert.Contains("Extension=\"png\"", ct);
            Assert.Contains("/xl/drawings/drawing1.xml", ct);

            // 不误报图片降级
            Assert.DoesNotContain(wb.SaveDegradations, d => d.Capability == DegradationCapability.Images);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Xlsb_FloatingImage_OpenSave_NoDuplicate()
    {
        // 打开再保存：图片不翻倍，drawing 合并而非重复。
        var file1 = GetTempFile(".xlsb");
        var file2 = GetTempFile(".xlsb");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "x");
            ws.AddImage(Png1x1, 2, 1, widthPx: 40, heightPx: 40, placement: ImagePlacement.Floating);
            wb.SaveAs(file1);

            var reopened = Excel.Open(file1);
            reopened.SaveAs(file2, ExcelFormat.Xlsb);

            int media1 = CountEntries(file1, "xl/media/image");
            int media2 = CountEntries(file2, "xl/media/image");
            Assert.Equal(media1, media2);

            var drawing2 = ReadEntry(file2, "xl/drawings/drawing1.xml");
            int anchors = System.Text.RegularExpressions.Regex.Matches(drawing2, "<xdr:pic>").Count;
            Assert.Equal(1, anchors);
        }
        finally { if (File.Exists(file1)) File.Delete(file1); if (File.Exists(file2)) File.Delete(file2); }
    }

    private static int CountEntries(string file, string prefix)
    {
        using var zip = ZipFile.OpenRead(file);
        return zip.Entries.Count(e => e.FullName.StartsWith(prefix, System.StringComparison.Ordinal));
    }

    /// <summary>解析 sheetN.bin 的 BIFF12 记录类型集合（用于断言 BrtDrawing 存在）。</summary>
    private static System.Collections.Generic.HashSet<int> RecordTypes(byte[] b)
    {
        var set = new System.Collections.Generic.HashSet<int>();
        int pos = 0;
        while (pos < b.Length)
        {
            int rt = ReadVarInt(b, ref pos);
            int len = ReadVarInt(b, ref pos);
            if (len < 0 || pos + len > b.Length) break;
            set.Add(rt);
            pos += len;
        }
        return set;
    }

    private static int ReadVarInt(byte[] b, ref int pos)
    {
        int v = 0, sh = 0;
        for (int i = 0; i < 4; i++)
        {
            byte x = b[pos++];
            v |= (x & 0x7F) << sh;
            if ((x & 0x80) == 0) return v;
            sh += 7;
        }
        return v;
    }
}
