using LiteExcel.Internal.Biff12;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace LiteExcel.Tests;

/// <summary>
/// XlsbDxfTranscoder：xlsb styles.bin 的 BrtDXF(0x01FB) → xlsx &lt;dxf&gt; 转码。
/// 块类型 → OOXML 子元素映射经真实 Excel 样本标定（pivot 引用的 7 类全覆盖）。
/// 注：完整 xf 表与 pivot &lt;formats&gt; 一并启用会触发 Excel 拒开（numFmtId/numFmt 校验未决），
/// 故默认门控关闭（LITEXCEL_ENABLE_PIVOT_FORMATS=1）；本测试仅验证解码逻辑。
/// </summary>
public class XlsbDxfTranscoderTests
{
    private const string Sample = @"D:\AiStory\Test\raw_repro.xlsb";

    private static byte[]? ReadStylesBin()
    {
        if (!File.Exists(Sample)) return null;
        using var zip = ZipFile.OpenRead(Sample);
        var e = zip.GetEntry("xl/styles.bin");
        if (e is null) return null;
        using var ms = new MemoryStream();
        using (var s = e.Open()) s.CopyTo(ms);
        return ms.ToArray();
    }

    [Fact]
    public void TranscodeAll_ProducesDxfPerBrtDXF_WhenRealFileAvailable()
    {
        var bin = ReadStylesBin();
        if (bin is null) return; // fixture 缺失时跳过（CI 无样本）

        var result = XlsbDxfTranscoder.TranscodeAll(bin);
        // IndexMap 覆盖全部原始 BrtDXF 序号（0..N-1），去重后内容数 ≤ 原始数。
        // 不断言具体条数：样本文件会被更新，硬编码计数会无谓地变脆。
        Assert.True(result.IndexMap.Count > 0);
        Assert.Equal(Enumerable.Range(0, result.IndexMap.Count), result.IndexMap.Keys.OrderBy(k => k));
        // 去重后内容数不超过原始数（Excel 同款去重）
        Assert.True(result.Dxfs.Count > 0 && result.Dxfs.Count <= result.IndexMap.Count);

        // 每个 dxf 均为 <dxf>...</dxf>
        Assert.All(result.Dxfs, d => Assert.StartsWith("<dxf>", d));
    }

    [Fact]
    public void TranscodeAll_FillBgColor_ThemeDecodedCorrectly()
    {
        var bin = ReadStylesBin();
        if (bin is null) return;

        var result = XlsbDxfTranscoder.TranscodeAll(bin);
        // 红色填充 dxf（BrtColor 05 FF 00 00 FF 00 00 FF → rgb="FFFF0000"）应解码出 solid + bgColor。
        // 不断言固定索引：样本更新会移动 dxf 序号。
        Assert.Contains(result.Dxfs, d =>
            d.Contains("<bgColor rgb=\"FFFF0000\"/>") && d.Contains("patternType=\"solid\""));
    }

    [Fact]
    public void TranscodeAll_FontName_UsesU16LengthPrefix()
    {
        var bin = ReadStylesBin();
        if (bin is null) return;

        var result = XlsbDxfTranscoder.TranscodeAll(bin);
        // dxf 块内字符串用 u16 前缀（非 XLWideString 的 u32）；字体名须正确解码为 "Book Antiqua"
        Assert.Contains(result.Dxfs, d => d.Contains("<name val=\"Book Antiqua\"/>"));
    }

    [Fact]
    public void TranscodeAll_RegistersCustomNumFmt()
    {
        var bin = ReadStylesBin();
        if (bin is null) return;

        var result = XlsbDxfTranscoder.TranscodeAll(bin);
        // 自定义 numFmtId(178) 须收集以供注册进 <numFmts>
        Assert.True(result.CustomNumFmts.ContainsKey(178));
        Assert.Contains("m/d/yyyy", result.CustomNumFmts[178]);
    }

    [Fact]
    public void TranscodeAll_FontBeforeNumFmt_InDxfChildOrder()
    {
        var bin = ReadStylesBin();
        if (bin is null) return;

        var result = XlsbDxfTranscoder.TranscodeAll(bin);
        // OOXML CT_Dxf 序要求 font 在 numFmt 之前；越序会被 Excel 判 schema 违规整包拒开。
        // 找含 <font> 与 <numFmt> 的 dxf，断言 font 先出现。
        foreach (var d in result.Dxfs)
        {
            int fi = d.IndexOf("<font>", StringComparison.Ordinal);
            int ni = d.IndexOf("<numFmt ", StringComparison.Ordinal);
            if (fi >= 0 && ni >= 0)
                Assert.True(fi < ni, $"dxf 中 font 应在 numFmt 之前: {d}");
        }
    }

    [Fact]
    public void BuildDxfsXml_WrapsInDxfsElementWithCount()
    {
        var bin = ReadStylesBin();
        if (bin is null) return;

        var result = XlsbDxfTranscoder.TranscodeAll(bin);
        var xml = XlsbDxfTranscoder.BuildDxfsXml(result.Dxfs);
        Assert.NotNull(xml);
        Assert.Contains($"<dxfs count=\"{result.Dxfs.Count}\">", xml);
        Assert.EndsWith("</dxfs>", xml);
    }

    [Fact]
    public void BuildDxfsXml_ReturnsNullForEmptyList()
    {
        Assert.Null(XlsbDxfTranscoder.BuildDxfsXml(new List<string>()));
        Assert.Null(XlsbDxfTranscoder.BuildDxfsXml(null!));
    }
}
