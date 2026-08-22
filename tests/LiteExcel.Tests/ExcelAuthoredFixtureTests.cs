using LiteExcel;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace LiteExcel.Tests;

/// <summary>
/// B1：真实 Excel（COM 生成）样本对拍。
/// 目的：验证库读取真实 Excel 产物的结果，而非仅"库自写自读"。
/// 覆盖：条件格式（cellIs/expression/colorScale/dataBar）、浮动图片、图表部件保留。
/// </summary>
public class ExcelAuthoredFixtureTests
{
    private static string Fixture(string name) =>
        Path.Combine(AppContext.BaseDirectory, "Fixtures", name);

    private static string ReadEntry(string file, string entry)
    {
        using var zip = ZipFile.OpenRead(file);
        var e = zip.GetEntry(entry);
        if (e is null) return "";
        using var s = e.Open();
        using var r = new StreamReader(s, Encoding.UTF8);
        return r.ReadToEnd();
    }

    [Fact]
    public void Read_ConditionalFormats_FromRealExcel()
    {
        // 真实 Excel 生成：A1:A10 cellIs>50 红加粗、B1:B10 expression 灰底、
        // C1:C10 colorScale 红→绿、D1:D10 dataBar
        var file = Fixture("excel-authored-cf.xlsx");
        Assert.True(File.Exists(file), $"missing fixture: {file}");

        var wb = Excel.Open(file);
        var ws = wb.Worksheets[0];
        var rules = ws.ConditionalFormats;
        Assert.NotNull(rules);
        Assert.True(rules.Count >= 4, $"expected >=4 rules, got {rules.Count}");
    }

    [Fact]
    public void Read_FloatingImage_FromRealExcel()
    {
        // 真实 Excel 生成：C2 附近 1 张浮动图片（64x64）
        var file = Fixture("excel-authored-image-chart.xlsx");
        Assert.True(File.Exists(file), $"missing fixture: {file}");

        var wb = Excel.Open(file);
        var ws = wb.Worksheets[0];
        Assert.NotEmpty(ws.Images);
        var img = ws.Images[0];
        Assert.Equal(ImagePlacement.Floating, img.Placement);
        Assert.True(img.Data.Length > 0, "image data empty");
        Assert.False(string.IsNullOrEmpty(img.Extension));
    }

    [Fact]
    public void RealExcel_ChartPart_PreservedOnSave()
    {
        // 真实 Excel 生成的图表：打开-保存后 chart 部件/关系保留
        var file = Fixture("excel-authored-image-chart.xlsx");
        Assert.True(File.Exists(file), $"missing fixture: {file}");

        var tmp = Path.Combine(Path.GetTempPath(), $"b1_{Guid.NewGuid():N}.xlsx");
        try
        {
            var wb = Excel.Open(file);
            wb.Worksheets[0].SetValue("B1", "改");
            wb.SaveAs(tmp);

            using (var zip = ZipFile.OpenRead(tmp))
            {
                Assert.NotNull(zip.GetEntry("xl/charts/chart1.xml"));
                Assert.NotNull(zip.GetEntry("xl/drawings/drawing1.xml"));
                Assert.NotNull(zip.GetEntry("xl/media/image1.png"));
            }
        }
        finally { if (File.Exists(tmp)) File.Delete(tmp); }
    }

    [Fact]
    public void RealExcel_ReadCellValues()
    {
        var file = Fixture("excel-authored-cf.xlsx");
        var wb = Excel.Open(file);
        var ws = wb.Worksheets[0];
        Assert.Equal(10, ws.Cell("A1").GetDouble());
        Assert.Equal(100, ws.Cell("A10").GetDouble());
    }
}
