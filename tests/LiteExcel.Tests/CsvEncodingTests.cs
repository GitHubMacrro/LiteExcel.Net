using System.Text;
using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// CSV 编码选项：ExcelReadOptions.Encoding / ExcelWriteOptions.Encoding。
/// 编码实例由调用方提供（库零依赖）；此处只用 BCL 自带编码验证透传链，
/// GBK(936) 在 net8.0 需调用方注册 CodePagesEncodingProvider，故不做自动化断言。
/// </summary>
public class CsvEncodingTests
{
    private static string Tmp() => Path.Combine(Path.GetTempPath(), $"csvenc_{Guid.NewGuid():N}.csv");

    private static Workbook MakeBook(string a1 = "中文测试", string b1 = "abc")
    {
        var wb = Excel.Create(ExcelFormat.Csv);
        var ws = wb.Worksheets[0];
        ws.SetValue("A1", a1);
        ws.SetValue("B1", b1);
        return wb;
    }

    [Fact]
    public void Default_WritesUtf8Bom()
    {
        var path = Tmp();
        try
        {
            Excel.Write(path, MakeBook());
            var bytes = File.ReadAllBytes(path);
            // 默认应写 UTF-8 BOM（EF BB BF）
            Assert.True(bytes.Length >= 3);
            Assert.Equal(0xEF, bytes[0]);
            Assert.Equal(0xBB, bytes[1]);
            Assert.Equal(0xBF, bytes[2]);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Default_RoundTrip_ChineseIntact()
    {
        var path = Tmp();
        try
        {
            Excel.Write(path, MakeBook("中文测试"));
            var wb = Excel.Open(path);
            Assert.Equal("中文测试", wb.Worksheets[0].Cell("A1").GetString());
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Utf8WithoutBom_NoPreambleWritten()
    {
        var path = Tmp();
        try
        {
            Excel.Write(path, MakeBook(), new ExcelWriteOptions
            {
                Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            });
            var bytes = File.ReadAllBytes(path);
            // 无 BOM：首字节应是内容而非 EF
            Assert.NotEqual(0xEF, bytes[0]);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Unicode_RoundTrip()
    {
        var path = Tmp();
        try
        {
            Excel.Write(path, MakeBook("中文测试"), new ExcelWriteOptions { Encoding = Encoding.Unicode });
            var wb = Excel.Open(path, new ExcelReadOptions { Encoding = Encoding.Unicode });
            Assert.Equal("中文测试", wb.Worksheets[0].Cell("A1").GetString());
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Utf32_RoundTrip()
    {
        var path = Tmp();
        try
        {
            Excel.Write(path, MakeBook("中文测试"), new ExcelWriteOptions { Encoding = Encoding.UTF32 });
            var wb = Excel.Open(path, new ExcelReadOptions { Encoding = Encoding.UTF32 });
            Assert.Equal("中文测试", wb.Worksheets[0].Cell("A1").GetString());
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Latin1_RoundTrip_AsciiIntact()
    {
        // Latin1(28591) 是 BCL 自带（AOT + InvariantGlobalization 下亦可用）
        var latin1 = Encoding.GetEncoding(28591);
        var path = Tmp();
        try
        {
            Excel.Write(path, MakeBook("Hello", "World"), new ExcelWriteOptions { Encoding = latin1 });
            var bytes = File.ReadAllBytes(path);
            Assert.NotEqual(0xEF, bytes[0]); // Latin1 无 preamble

            var wb = Excel.Open(path, new ExcelReadOptions { Encoding = latin1 });
            Assert.Equal("Hello", wb.Worksheets[0].Cell("A1").GetString());
            Assert.Equal("World", wb.Worksheets[0].Cell("B1").GetString());
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void MismatchedEncoding_ProducesGarbledText()
    {
        // 写 UTF-32、按 Latin1 读 → 中文必然错乱（证明参数真的生效，而非被忽略）
        var path = Tmp();
        try
        {
            Excel.Write(path, MakeBook("中文测试"), new ExcelWriteOptions { Encoding = Encoding.UTF32 });
            var wb = Excel.Open(path, new ExcelReadOptions { Encoding = Encoding.GetEncoding(28591) });
            Assert.NotEqual("中文测试", wb.Worksheets[0].Cell("A1").GetString());
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void ExplicitEncoding_TakesPrecedenceOverBom()
    {
        // 文件带 UTF-8 BOM，但显式指定 Latin1 读 → 应按 Latin1 解码（BOM 字节被当作 Latin1 字符）
        var path = Tmp();
        try
        {
            Excel.Write(path, MakeBook("abc", "def")); // 默认写 UTF-8 + BOM
            var wb = Excel.Open(path, new ExcelReadOptions { Encoding = Encoding.GetEncoding(28591) });
            var a1 = wb.Worksheets[0].Cell("A1").GetString() ?? "";
            // Latin1 解码时 BOM 三字节变成可见字符，A1 不再等于纯 "abc"
            Assert.NotEqual("abc", a1);
            Assert.EndsWith("abc", a1);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void NoExplicitEncoding_BomDetectionStillWorks()
    {
        // 未指定 Encoding 时，BOM 探测保持生效
        var path = Tmp();
        try
        {
            Excel.Write(path, MakeBook("中文测试")); // UTF-8 + BOM
            var wb = Excel.Open(path); // 不传 Encoding
            Assert.Equal("中文测试", wb.Worksheets[0].Cell("A1").GetString());
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void EncodingWithSeparator_BothApplied()
    {
        var path = Tmp();
        try
        {
            Excel.Write(path, MakeBook("中文", "值"), new ExcelWriteOptions
            {
                Encoding = Encoding.Unicode,
                Separator = ';',
            });
            var wb = Excel.Open(path, new ExcelReadOptions
            {
                Encoding = Encoding.Unicode,
                Separator = ';',
            });
            Assert.Equal("中文", wb.Worksheets[0].Cell("A1").GetString());
            Assert.Equal("值", wb.Worksheets[0].Cell("B1").GetString());
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void GbkRoundTrip_WhenProviderAvailable()
    {
        // GBK(936)：net48 BCL 自带；net8.0 需调用方注册 CodePagesEncodingProvider。
        // 不可用时跳过（库本身不引用编码包，测试项目同样保持零依赖）。
        Encoding gbk;
        try { gbk = Encoding.GetEncoding(936); }
        catch { return; }

        var path = Tmp();
        try
        {
            Excel.Write(path, MakeBook("中文测试"), new ExcelWriteOptions { Encoding = gbk });
            var bytes = File.ReadAllBytes(path);
            Assert.NotEqual(0xEF, bytes[0]); // GBK 无 preamble

            var wb = Excel.Open(path, new ExcelReadOptions { Encoding = gbk });
            Assert.Equal("中文测试", wb.Worksheets[0].Cell("A1").GetString());
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
