using LiteExcel;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace LiteExcel.Tests;

/// <summary>
/// P1-B：浮动图片读回（oneCellAnchor / twoCellAnchor / editAs=absolute）。
/// 用库写出的文件验证读侧解析，不与真实 Excel 样本一拍。
/// </summary>
public class ImageReadBackTests
{
    private static readonly byte[] Png1x1 = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");

    private static string GetTempFile() =>
        Path.Combine(Path.GetTempPath(), $"p1b_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void WriteAndReadBack_OneCellAnchor()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].AddImage(Png1x1, new ImageAnchor
            {
                TopLeftCell = "B3",
                WidthPixels = 64,
                HeightPixels = 64,
            }, extension: "png", name: "Logo", altText: "公司 Logo");
            wb.SaveAs(file);

            var rb = Excel.Open(file);
            var imgs = rb.Worksheets[0].Images;
            Assert.Single(imgs);
            var img = imgs[0];
            Assert.Equal(ImagePlacement.Floating, img.Placement);
            Assert.Equal("Logo", img.Name);
            Assert.Equal("公司 Logo", img.AltText);
            Assert.Equal(3, img.Row);
            Assert.Equal(2, img.Column);
            Assert.Equal("png", img.Extension);
            Assert.True(img.Data.Length > 0);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void ReadBack_TwoCellAnchor_WithOffsets()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].AddImage(Png1x1, new ImageAnchor
            {
                TopLeftCell = "C4",
                TopLeftOffsetX = 9525,
                TopLeftOffsetY = 19050,
                WidthPixels = 80,
                HeightPixels = 60,
                MoveMode = ImageMoveMode.MoveAndSizeWithCells,
            }, extension: "png", name: "Shape");
            wb.SaveAs(file);

            var rb = Excel.Open(file);
            var img = rb.Worksheets[0].Images[0];
            Assert.Equal("C4", img.Anchor!.TopLeftCell);
            Assert.Equal(9525, img.Anchor.TopLeftOffsetX);
            Assert.Equal(19050, img.Anchor.TopLeftOffsetY);
            Assert.Equal(ImageMoveMode.MoveAndSizeWithCells, img.Anchor.MoveMode);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void ReadBack_FixedPosition()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].AddImage(Png1x1, new ImageAnchor
            {
                TopLeftCell = "D2",
                WidthPixels = 50,
                HeightPixels = 50,
                MoveMode = ImageMoveMode.FixedPosition,
            }, extension: "png", name: "Pin");
            wb.SaveAs(file);

            var rb = Excel.Open(file);
            var img = rb.Worksheets[0].Images[0];
            Assert.Equal(ImageMoveMode.FixedPosition, img.Anchor!.MoveMode);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void ReadBack_MixedSheetOnlyFloatingInFloatingSheet()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].SetValue("A1", "no");
            wb.Worksheets[0].AddImage(Png1x1, 2, 1, 32, 32, ImagePlacement.Floating, "png");
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets["Sheet2"].AddImage(Png1x1, 5, 3, 32, 32, ImagePlacement.Floating, "png");
            wb.SaveAs(file);

            var rb = Excel.Open(file);
            Assert.Single(rb.Worksheets["Sheet1"].Images);
            Assert.Single(rb.Worksheets["Sheet2"].Images);
            Assert.Equal(2, rb.Worksheets["Sheet1"].Images[0].Row);
            Assert.Equal(5, rb.Worksheets["Sheet2"].Images[0].Row);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void ReadBack_DataBytesRoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].AddImage(Png1x1, 1, 1, 32, 32, ImagePlacement.Floating, "png", "r1");
            wb.SaveAs(file);

            var rb = Excel.Open(file);
            Assert.True(rb.Worksheets[0].Images[0].Data.Length > 0);

            using var zip = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            var media = zip.GetEntry("xl/media/image1.png");
            Assert.NotNull(media);
            byte[] raw;
            using (var ms = new MemoryStream())
            {
                using var s = media!.Open();
                s.CopyTo(ms);
                raw = ms.ToArray();
            }
            Assert.Equal(raw, rb.Worksheets[0].Images[0].Data);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
