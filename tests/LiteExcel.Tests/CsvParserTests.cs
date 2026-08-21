using LiteExcel;
using System.Text;

namespace LiteExcel.Tests;

public class CsvParserTests
{
    private static Workbook Read(string csv)
    {
        using var stream = new MemoryStream(new UTF8Encoding(false).GetBytes(csv));
        return Excel.Open(stream, ExcelFormat.Csv);
    }

    [Fact]
    public void QuotedNewline_IsKeptInsideOneField()
    {
        var wb = Read("a,b\n\"first line\nsecond line\",2\n");
        var sheet = wb.Worksheets[0].ToSheetData();

        Assert.Equal(2, sheet.Rows.Count);
        Assert.Equal("first line\nsecond line", sheet.Rows[1][0].GetString());
        Assert.Equal("2", sheet.Rows[1][1].GetString());
    }

    [Fact]
    public void EmptyRows_ArePreserved()
    {
        var wb = Read("a,b\n\n1,2\n\n");
        var rows = wb.Worksheets[0].ToSheetData().Rows;

        Assert.Equal(4, rows.Count);
        Assert.Single(rows[1]);
        Assert.All(rows[1], cell => Assert.True(cell.IsEmpty));
        Assert.Equal("1", rows[2][0].GetString());
        Assert.All(rows[3], cell => Assert.True(cell.IsEmpty));
    }

    [Fact]
    public void EscapedQuotesAndEmptyFields_RoundTrip()
    {
        var wb = Read("\"a\"\"b\",,c\r\n");
        var row = wb.Worksheets[0].ToSheetData().Rows[0];

        Assert.Equal(3, row.Count);
        Assert.Equal("a\"b", row[0].GetString());
        Assert.True(row[1].IsEmpty);
        Assert.Equal("c", row[2].GetString());
    }

    [Fact]
    public void EmptyFile_HasNoRows()
    {
        var wb = Read("");
        Assert.Empty(wb.Worksheets[0].ToSheetData().Rows);
    }

    [Fact]
    public void UnterminatedQuotedField_ThrowsFormatException()
    {
        Assert.Throws<FormatException>(() => Read("a,\"unterminated\n"));
    }
}
