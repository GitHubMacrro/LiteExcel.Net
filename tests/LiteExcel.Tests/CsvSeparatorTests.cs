using LiteExcel;
using System.Text;

namespace LiteExcel.Tests;

public class CsvSeparatorTests
{
    private static string GetTempFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"p1a_{Guid.NewGuid():N}{ext}");

    [Fact]
    public void AutoDetect_CommaOnDefault()
    {
        var file = GetTempFile(".csv");
        File.WriteAllText(file, "a,b\n1,2\n");
        try
        {
            var wb = Excel.Open(file);
            var sheet = wb.Worksheets[0].ToSheetData();
            Assert.Equal(2, sheet.Rows[0].Count);
            Assert.Equal("a", sheet.Rows[0][0].GetString());
            Assert.Equal("b", sheet.Rows[0][1].GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void AutoDetect_Semicolon_WinsByFrequency()
    {
        var file = GetTempFile(".csv");
        File.WriteAllText(file, "a;b;c\n1;2;3\n4;5;6\n");
        try
        {
            var wb = Excel.Open(file);
            var sheet = wb.Worksheets[0].ToSheetData();
            Assert.Equal(3, sheet.Rows[0].Count);
            Assert.Equal("a", sheet.Rows[0][0].GetString());
            Assert.Equal("c", sheet.Rows[0][2].GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void AutoDetect_Tab_Wins()
    {
        var file = GetTempFile(".csv");
        File.WriteAllText(file, "h1\th2\th3\n11\t22\t33\n");
        try
        {
            var wb = Excel.Open(file);
            var sheet = wb.Worksheets[0].ToSheetData();
            Assert.Equal(3, sheet.Rows[0].Count);
            Assert.Equal("h2", sheet.Rows[0][1].GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void ExplicitSeparator_OverridesAutoDetect()
    {
        var file = GetTempFile(".csv");
        File.WriteAllText(file, "a;b\n1;2\n");
        try
        {
            // 强制逗号 → 分号被当成普通字符
            var wb = Excel.Open(file, new ExcelReadOptions { Separator = ',' });
            var sheet = wb.Worksheets[0].ToSheetData();
            Assert.Single(sheet.Rows[0]);
            Assert.Equal("a;b", sheet.Rows[0][0].GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void QuotedSeparator_NotCounted()
    {
        var file = GetTempFile(".csv");
        File.WriteAllText(file, "a,\"x,y\",c\n");
        try
        {
            var wb = Excel.Open(file);
            var sheet = wb.Worksheets[0].ToSheetData();
            Assert.Equal(3, sheet.Rows[0].Count);
            Assert.Equal("x,y", sheet.Rows[0][1].GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Write_WithSemicolonSeparator()
    {
        var file = GetTempFile(".csv");
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].SetValue("A1", "a");
            wb.Worksheets[0].SetValue("B1", "b");
            wb.Worksheets[0].SetValue("C1", "c;d"); //含分号
            Excel.Write(file, wb, new ExcelWriteOptions { Separator = ';' });

            var text = File.ReadAllText(file);
            Assert.Contains("a;b", text);
            Assert.Contains("\"c;d\"", text); // 含分号被引号包裹
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void ReadAndWrite_TabRoundTrip()
    {
        var file = GetTempFile(".csv");
        File.WriteAllText(file, "h1\th2\n1\t2\n");
        try
        {
            var wb1 = Excel.Open(file, new ExcelReadOptions { Separator = '\t' });
            var sheet1 = wb1.Worksheets[0].ToSheetData();
            Assert.Equal("h2", sheet1.Rows[0][1].GetString());

            wb1.Worksheets[0].SetValue("A2", "x");
            Excel.Write(file, wb1, new ExcelWriteOptions { Separator = '\t' });

            var wb2 = Excel.Open(file, new ExcelReadOptions { Separator = '\t' });
            var sheet2 = wb2.Worksheets[0].ToSheetData();
            Assert.Equal("h2", sheet2.Rows[0][1].GetString());
            Assert.Equal("x", sheet2.Rows[1][0].GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
