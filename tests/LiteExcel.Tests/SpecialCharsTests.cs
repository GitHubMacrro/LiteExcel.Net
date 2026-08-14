using LiteExcel;

namespace LiteExcel.Tests;

public class SpecialCharsTests
{
    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"litexlsx_{Guid.NewGuid():N}.xlsx");

    [Theory]
    [InlineData("&", "ampersand")]
    [InlineData("<", "less than")]
    [InlineData(">", "greater than")]
    [InlineData("\"", "quote")]
    [InlineData("'", "apostrophe")]
    [InlineData("&<>\"'", "all special")]
    public void XmlSpecialChars_RoundTrip(string text, string label)
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                Headers = new() { "col" },
                Rows = new() { new Cell[] { Cell.FromText(text) } },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Equal(text, read.Rows[0][0].Text);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void ChineseChars_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var texts = new[] { "你好世界", "中文测试ＡＢＣ", "繁體字測試", "全角符号！＠＃￥％" };
            var sheet = new SheetData
            {
                Headers = new() { "中文" },
                Rows = texts.Select(t => (IReadOnlyList<Cell>)new Cell[] { Cell.FromText(t) }).ToList(),
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            for (int i = 0; i < texts.Length; i++)
            {
                Assert.Equal(texts[i], read.Rows[i][0].Text);
            }
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Emoji_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var text = "emoji 😀🎉 test";
            var sheet = new SheetData
            {
                Headers = new() { "emoji" },
                Rows = new() { new Cell[] { Cell.FromText(text) } },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Equal(text, read.Rows[0][0].Text);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void LeadingTrailingSpaces_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var text = "  spaces  ";
            var sheet = new SheetData
            {
                Headers = new() { "col" },
                Rows = new() { new Cell[] { Cell.FromText(text) } },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Equal(text, read.Rows[0][0].Text);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void LongText_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var text = new string('A', 10000);
            var sheet = new SheetData
            {
                Headers = new() { "long" },
                Rows = new() { new Cell[] { Cell.FromText(text) } },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Equal(text, read.Rows[0][0].Text);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
