using LiteExcel;
using System.IO;

namespace LiteExcel.Tests;

/// <summary>
/// A1：InCell richData 图片读回。
/// 库写 InCell 图 → Excel.Open → Worksheet.Images 应包含 InCell 条目（Data/Row/Column/Extension）。
/// </summary>
public class InCellImageReadBackTests
{
    private static readonly byte[] Png1x1 = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");

    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"a1_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void ReadBack_InCellImage_FromWrittenFile()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].Cell("A1").SetValue("InCell");
            wb.Worksheets[0].AddImage(Png1x1, 2, 1, placement: ImagePlacement.InCell, extension: "png", name: "ic");
            wb.SaveAs(file);

            var rb = Excel.Open(file);
            var imgs = rb.Worksheets[0].Images;
            Assert.NotEmpty(imgs);
            var inCell = imgs.FirstOrDefault(i => i.Placement == ImagePlacement.InCell);
            Assert.NotNull(inCell);
            Assert.Equal(2, inCell!.Row);
            Assert.Equal(1, inCell.Column);
            Assert.Equal("png", inCell.Extension);
            Assert.True(inCell.Data.Length > 0);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void ReadBack_MixedFloatingAndInCell()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].AddImage(Png1x1, 1, 1, 32, 32, ImagePlacement.Floating, "png", "f1");
            wb.Worksheets[0].AddImage(Png1x1, 5, 2, placement: ImagePlacement.InCell, extension: "png", name: "i2");
            wb.SaveAs(file);

            var rb = Excel.Open(file);
            var imgs = rb.Worksheets[0].Images;
            Assert.Equal(2, imgs.Count);
            Assert.Contains(imgs, i => i.Placement == ImagePlacement.Floating);
            Assert.Contains(imgs, i => i.Placement == ImagePlacement.InCell);
            var inCell = imgs.First(i => i.Placement == ImagePlacement.InCell);
            Assert.Equal(5, inCell.Row);
            Assert.Equal(2, inCell.Column);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void ReadBack_MultiSheetInCell_GlobalVmIndexing()
    {
        // vm 是工作簿级全局索引：sheet1 + sheet2 各 1 个 InCell，读回需都回填到正确 sheet
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].AddImage(Png1x1, 2, 1, placement: ImagePlacement.InCell, extension: "png", name: "s1");
            var ws2 = wb.Worksheets.Add("S2");
            ws2.AddImage(Png1x1, 3, 4, placement: ImagePlacement.InCell, extension: "png", name: "s2");
            wb.SaveAs(file);

            var rb = Excel.Open(file);
            Assert.Single(rb.Worksheets[0].Images.Where(i => i.Placement == ImagePlacement.InCell));
            Assert.Single(rb.Worksheets[1].Images.Where(i => i.Placement == ImagePlacement.InCell));
            Assert.Equal(2, rb.Worksheets[0].Images.First().Row);
            Assert.Equal(3, rb.Worksheets[1].Images.First().Row);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
