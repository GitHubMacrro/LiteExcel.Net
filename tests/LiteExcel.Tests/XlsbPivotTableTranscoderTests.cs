using LiteExcel.Internal.Biff12;

namespace LiteExcel.Tests;

/// <summary>
/// Stage D（进行中）：pivotTable BIFF12 → 模型/XML 的核心段解析（头 / location / 字段数 / 样式）。
/// 字段与项结构（62 种记录类型）待补。
/// </summary>
public class XlsbPivotTableTranscoderTests
{
    [Fact]
    public void PivotTable_ParsesHeaderLocationAndStyle_WhenRealFileAvailable()
    {
        var path = @"D:\AiStory\Test\raw_repro.xlsb";
        if (!File.Exists(path)) return;
        using var zip = System.IO.Compression.ZipFile.OpenRead(path);
        var e = zip.GetEntry("xl/pivotTables/pivotTable4.bin");
        if (e is null) return;
        using var ms = new MemoryStream();
        using (var s = e.Open()) s.CopyTo(ms);

        var info = XlsbPivotTableTranscoder.Parse(ms.ToArray());
        Assert.Equal(1u, info.CacheId);
        Assert.Equal("PivotTable1", info.Name);
        Assert.Equal("Values", info.DataCaption);
        Assert.True(info.HasLocation);
        Assert.Equal((0, 64, 3, 3), info.Location);
        Assert.Equal(1, info.FirstHeaderRow);
        Assert.Equal(1, info.FirstDataRow);
        Assert.Equal(1, info.FirstDataCol);
        Assert.Equal(1, info.PivotFieldCount);
        Assert.Equal("PivotStyleLight16", info.PivotStyle);

        var xml = XlsbPivotTableTranscoder.ToXml(info);
        Assert.Contains("name=\"PivotTable1\"", xml);
        Assert.Contains("cacheId=\"1\"", xml);
        Assert.Contains("dataCaption=\"Values\"", xml);
        Assert.Contains("ref=\"D1:D65\"", xml);
    }
}
