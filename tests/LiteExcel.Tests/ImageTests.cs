using LiteExcel;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace LiteExcel.Tests;

/// <summary>
/// Phase 7：图片写回测试（Floating drawing + InCell richData）。
/// </summary>
public class ImageTests
{
    // 1x1 有效 PNG（透明）
    private static readonly byte[] Png1x1 = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");

    private static string Save(Workbook wb)
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        wb.SaveAs(path);
        return path;
    }

    private static string ReadEntry(string xlsx, string entry)
    {
        using var zip = ZipFile.OpenRead(xlsx);
        var e = zip.GetEntry(entry);
        if (e is null) return "";
        using var r = new StreamReader(e.Open());
        return r.ReadToEnd();
    }

    private static byte[] ReadEntryBytes(string xlsx, string entry)
    {
        using var zip = ZipFile.OpenRead(xlsx);
        var e = zip.GetEntry(entry);
        if (e is null) return Array.Empty<byte>();
        using var ms = new MemoryStream();
        using (var s = e.Open()) s.CopyTo(ms);
        return ms.ToArray();
    }

    [Fact]
    public void FloatingImage_WritesMediaAndDrawing()
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("图片");
        wb.Worksheets[0].AddImage(Png1x1, 2, 1, widthPx: 50, heightPx: 50, placement: ImagePlacement.Floating);

        var path = Save(wb);
        try
        {
            Assert.True(File.Exists(path));
            // media 文件写入
            var media = ReadEntryBytes(path, "xl/media/image1.png");
            Assert.Equal(Png1x1, media);
            // drawing 生成
            var drawing = ReadEntry(path, "xl/drawings/drawing1.xml");
            Assert.Contains("<xdr:oneCellAnchor>", drawing);
            Assert.Contains("r:embed=\"rId1\"", drawing);
            Assert.Contains("<xdr:col>0</xdr:col>", drawing);
            Assert.Contains("<xdr:row>1</xdr:row>", drawing);
            // sheet 引用 drawing
            var sheet1 = ReadEntry(path, "xl/worksheets/sheet1.xml");
            Assert.Contains("<drawing r:id=\"rIdD1\"/>", sheet1);
            // sheet rels 有 drawing 关系
            var sheetRels = ReadEntry(path, "xl/worksheets/_rels/sheet1.xml.rels");
            Assert.Contains("/drawing\" Target=\"../drawings/drawing1.xml\"", sheetRels);
            // drawing rels 指向 media
            var drawingRels = ReadEntry(path, "xl/drawings/_rels/drawing1.xml.rels");
            Assert.Contains("../media/image1.png", drawingRels);
            // Content_Types 有 png Default 与 drawing Override
            var ct = ReadEntry(path, "[Content_Types].xml");
            Assert.Contains("Extension=\"png\" ContentType=\"image/png\"", ct);
            Assert.Contains("/xl/drawings/drawing1.xml\"", ct);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void InCellImage_WritesRichDataParts()
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("InCell");
        wb.Worksheets[0].AddImage(Png1x1, 2, 1, placement: ImagePlacement.InCell);

        var path = Save(wb);
        try
        {
            // media 文件
            var media = ReadEntryBytes(path, "xl/media/image1.png");
            Assert.Equal(Png1x1, media);
            // sheet 单元格 t="e" vm="1"
            var sheet1 = ReadEntry(path, "xl/worksheets/sheet1.xml");
            Assert.Contains("<c r=\"A2\" t=\"e\" vm=\"1\"><v>#VALUE!</v></c>", sheet1);
            // metadata
            var metadata = ReadEntry(path, "xl/metadata.xml");
            Assert.Contains("futureMetadata name=\"XLRICHVALUE\"", metadata);
            Assert.Contains("valueMetadata count=\"1\"", metadata);
            // richValueRel 与 rels
            var richValueRel = ReadEntry(path, "xl/richData/richValueRel.xml");
            Assert.Contains("<rel r:id=\"rId1\"/>", richValueRel);
            var richValueRels = ReadEntry(path, "xl/richData/_rels/richValueRel.xml.rels");
            Assert.Contains("../media/image1.png", richValueRels);
            // rdrichvalue / structure / types
            Assert.Contains("<rvData", ReadEntry(path, "xl/richData/rdrichvalue.xml"));
            Assert.Contains("_localImage", ReadEntry(path, "xl/richData/rdrichvaluestructure.xml"));
            Assert.Contains("rvTypesInfo", ReadEntry(path, "xl/richData/rdRichValueTypes.xml"));
            // workbook rels 有 richData 关系
            var wbRels = ReadEntry(path, "xl/_rels/workbook.xml.rels");
            Assert.Contains("sheetMetadata\" Target=\"metadata.xml\"", wbRels);
            Assert.Contains("rdRichValue\" Target=\"richData/rdrichvalue.xml\"", wbRels);
            Assert.Contains("richValueRel\" Target=\"richData/richValueRel.xml\"", wbRels);
            // Content_Types
            var ct = ReadEntry(path, "[Content_Types].xml");
            Assert.Contains("/xl/metadata.xml\"", ct);
            Assert.Contains("/xl/richData/richValueRel.xml\"", ct);
            Assert.Contains("/xl/richData/rdrichvalue.xml\"", ct);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void MultipleInCellImages_IndexesSequential()
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("x");
        wb.Worksheets[0].AddImage(Png1x1, 2, 1, placement: ImagePlacement.InCell);
        wb.Worksheets[0].AddImage(Png1x1, 3, 2, placement: ImagePlacement.InCell);

        var path = Save(wb);
        try
        {
            var sheet1 = ReadEntry(path, "xl/worksheets/sheet1.xml");
            Assert.Contains("<c r=\"A2\" t=\"e\" vm=\"1\"", sheet1);
            Assert.Contains("<c r=\"B3\" t=\"e\" vm=\"2\"", sheet1);
            var metadata = ReadEntry(path, "xl/metadata.xml");
            Assert.Contains("valueMetadata count=\"2\"", metadata);
            Assert.Contains("<rc t=\"1\" v=\"0\"/>", metadata);
            Assert.Contains("<rc t=\"1\" v=\"1\"/>", metadata);
            var richValueRel = ReadEntry(path, "xl/richData/richValueRel.xml");
            Assert.Contains("<rel r:id=\"rId1\"/>", richValueRel);
            Assert.Contains("<rel r:id=\"rId2\"/>", richValueRel);
            var rvData = ReadEntry(path, "xl/richData/rdrichvalue.xml");
            Assert.Contains("count=\"2\"", rvData);
            var richValueRels = ReadEntry(path, "xl/richData/_rels/richValueRel.xml.rels");
            Assert.Contains("../media/image1.png", richValueRels);
            Assert.Contains("../media/image2.png", richValueRels);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void MixedImages_MultiSheet()
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("s1");
        wb.Worksheets[0].AddImage(Png1x1, 2, 1, widthPx: 30, heightPx: 30);
        var ws2 = wb.Worksheets.Add("Sheet2");
        ws2.Cell("A1").SetValue("s2");
        ws2.AddImage(Png1x1, 2, 1, placement: ImagePlacement.InCell);

        var path = Save(wb);
        try
        {
            // media 全局递增：sheet1 floating = image1, sheet2 incell = image2
            var sheet1Rels = ReadEntry(path, "xl/worksheets/_rels/sheet1.xml.rels");
            Assert.Contains("../drawings/drawing1.xml", sheet1Rels);
            Assert.Contains("<xdr:oneCellAnchor>", ReadEntry(path, "xl/drawings/drawing1.xml"));
            var richValueRels = ReadEntry(path, "xl/richData/_rels/richValueRel.xml.rels");
            Assert.Contains("../media/image2.png", richValueRels);
            // sheet2 无 drawing（只有 InCell）
            Assert.Equal("", ReadEntry(path, "xl/drawings/drawing2.xml"));
            var ct = ReadEntry(path, "[Content_Types].xml");
            Assert.Contains("/xl/drawings/drawing1.xml\"", ct);
            Assert.Contains("/xl/metadata.xml\"", ct);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void ImagePixelSize_ParsedFromHeader()
    {
        var img = new WorksheetImage { Data = Png1x1 };
        Assert.Equal((1, 1), img.PixelSize);
        Assert.Equal("png", img.EffectiveExtension);
    }

    // ── Phase 3 补：打开已有图片 → 追加图片 → 保存（zip 重名回归） ──

    [Fact]
    public void OpenFileWithImages_AddImage_NoDuplicateZipEntries()
    {
        // 先建一个带浮动图片的文件
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("Data");
        wb.Worksheets[0].AddImage(Png1x1, 2, 1, widthPx: 50, heightPx: 50, placement: ImagePlacement.Floating);
        var src = Save(wb);
        try
        {
            // 打开 → 追加第二张浮动图片 → 保存
            var opened = Excel.Open(src);
            opened.Worksheets[0].AddImage(Png1x1, 5, 3, widthPx: 40, heightPx: 40, placement: ImagePlacement.Floating);
            var outPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                opened.SaveAs(outPath);

                // 不应有 zip 重复条目；新 media 序号避开既有 image1.png
                using var zip = ZipFile.OpenRead(outPath);
                var names = zip.Entries.Select(e => e.FullName).ToList();
                Assert.Equal(names.Count, names.Distinct().Count());

                // 两张图片均在 drawing 中（原图保留 + 新图）
                var drawing = ReadEntry(outPath, "xl/drawings/drawing1.xml");
                Assert.Contains("<xdr:oneCellAnchor>", drawing);
                Assert.Contains("r:embed=\"rId1\"", drawing);
                Assert.Contains("r:embed=\"rId2\"", drawing);

                // 重新打开可读
                var reopened = Excel.Open(outPath);
                Assert.NotNull(reopened);
            }
            finally
            {
                if (File.Exists(outPath)) File.Delete(outPath);
            }
        }
        finally
        {
            if (File.Exists(src)) File.Delete(src);
        }
    }

    [Fact]
    public void OpenFileWithImages_AddImage_SameSheet_MergesDrawing()
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("Data");
        wb.Worksheets[0].AddImage(Png1x1, 2, 1, widthPx: 50, heightPx: 50, placement: ImagePlacement.Floating);
        var src = Save(wb);
        try
        {
            var opened = Excel.Open(src);
            // 同一 sheet 再加一张浮动图片 → 合并进既有 drawing1.xml
            opened.Worksheets[0].AddImage(Png1x1, 8, 8, widthPx: 30, heightPx: 30, placement: ImagePlacement.Floating);
            var outPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                opened.SaveAs(outPath);

                var drawing = ReadEntry(outPath, "xl/drawings/drawing1.xml");
                Assert.Contains("<xdr:oneCellAnchor>", drawing);
                // 原图 rId1 与新图 rId2
                Assert.Contains("r:embed=\"rId1\"", drawing);
                Assert.Contains("r:embed=\"rId2\"", drawing);

                var sheetRels = ReadEntry(outPath, "xl/worksheets/_rels/sheet1.xml.rels");
                // 仅一个 drawing 关系（不新增 rIdD1 副本）
                var count = System.Text.RegularExpressions.Regex.Matches(sheetRels, "/drawing\"").Count;
                Assert.Equal(1, count);

                // drawing rels 含两个 image 关系
                var drawingRels = ReadEntry(outPath, "xl/drawings/_rels/drawing1.xml.rels");
                Assert.Contains("../media/image1.png", drawingRels);
                Assert.Contains("../media/image2.png", drawingRels);
            }
            finally
            {
                if (File.Exists(outPath)) File.Delete(outPath);
            }
        }
        finally
        {
            if (File.Exists(src)) File.Delete(src);
        }
    }
}
