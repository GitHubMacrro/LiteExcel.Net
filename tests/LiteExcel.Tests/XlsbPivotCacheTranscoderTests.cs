using LiteExcel.Internal.Biff12;

namespace LiteExcel.Tests;

/// <summary>
/// Stage D（进行中）：pivotCacheDefinition BIFF12 → 模型/XML 的核心段解析
/// （头 / cacheSource / cacheFields / sharedItems）。cacheHierarchies 待补。
/// </summary>
public class XlsbPivotCacheTranscoderTests
{
    [Fact]
    public void PivotCache_ParsesHeaderSourceFieldsAndSharedItems()
    {
        var bin = XlsbTestFile.BuildPivotCacheBin(1, 12, "Admin", new[]
        {
            (2, 1, 0u, "[T].[A]", "A", new[] { "x", "y", "z" }),
            (3, 1, 0u, "[T].[B]", "B", Array.Empty<string>()),
        });

        var info = XlsbPivotCacheTranscoder.Parse(bin);
        Assert.Equal(8, info.RefreshedVersion);
        Assert.Equal(3, info.MinRefreshableVersion);
        Assert.Equal(5, info.CreatedVersion);
        Assert.Equal("Admin", info.RefreshedBy);
        Assert.Equal(1, info.SourceType);
        Assert.Equal(12u, info.ConnectionId);

        Assert.Equal(2, info.Fields.Count);
        Assert.Equal("[T].[A]", info.Fields[0].Name);
        Assert.Equal("A", info.Fields[0].Caption);
        Assert.Equal(2, info.Fields[0].Hierarchy);
        Assert.Equal(new[] { "x", "y", "z" }, info.Fields[0].SharedStrings);
        Assert.Equal("[T].[B]", info.Fields[1].Name);
        Assert.Empty(info.Fields[1].SharedStrings);

        var xml = XlsbPivotCacheTranscoder.ToXml(info);
        Assert.Contains("pivotCacheDefinition", xml);
        Assert.Contains("<cacheSource type=\"external\">", xml);
        Assert.Contains("cacheFields count=\"2\"", xml);
        Assert.Contains("name=\"[T].[A]\"", xml);
        Assert.Contains("caption=\"A\"", xml);
        Assert.Contains("<s v=\"x\"/>", xml);
    }

    [Fact]
    public void PivotCache_ParsesRealFile_WhenAvailable()
    {
        // 本机真实样本（可选；CI 无此文件时跳过）。
        var path = @"D:\AiStory\Test\raw_repro.xlsb";
        if (!File.Exists(path)) return;
        using var zip = System.IO.Compression.ZipFile.OpenRead(path);
        var e = zip.GetEntry("xl/pivotCache/pivotCacheDefinition4.bin");
        if (e is null) return;
        using var ms = new MemoryStream();
        using (var s = e.Open()) s.CopyTo(ms);

        var info = XlsbPivotCacheTranscoder.Parse(ms.ToArray());
        Assert.Equal(8, info.RefreshedVersion);
        Assert.Equal("Admin", info.RefreshedBy);
        Assert.Equal(1, info.SourceType);
        Assert.Equal(12u, info.ConnectionId);
        Assert.True(info.Fields.Count >= 12);
        Assert.Equal("PACKAGEGROUP", info.Fields[0].Caption);
        Assert.NotEmpty(info.Fields[0].SharedStrings);
    }
}
