using LiteExcel;

namespace LiteExcel.Tests;

public class RoundTripTests
{
    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"litexlsx_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void TextOnly_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "Test",
                Headers = new() { "A", "B", "C" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("hello"), Cell.FromText("world"), Cell.FromText("中文测试") },
                    new Cell[] { Cell.FromText(""), Cell.FromText("tab\there"), Cell.FromText("newline\nhere") },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Equal("Test", read.SheetName);
            Assert.Equal(3, read.Headers.Count);
            Assert.Equal("A", read.Headers[0]);
            Assert.Equal("B", read.Headers[1]);
            Assert.Equal("C", read.Headers[2]);

            Assert.Equal(2, read.Rows.Count);
            Assert.Equal("hello", read.Rows[0][0].Text);
            Assert.Equal("world", read.Rows[0][1].Text);
            Assert.Equal("中文测试", read.Rows[0][2].Text);

            // 空字符串读回应为 Empty
            Assert.Equal(CellType.Empty, read.Rows[1][0].Type);
            Assert.Equal("tab\there", read.Rows[1][1].Text);
            Assert.Equal("newline\nhere", read.Rows[1][2].Text);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Numbers_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                Headers = new() { "整数", "小数", "大数", "零", "负数" },
                Rows = new()
                {
                    new Cell[] { Cell.FromNumber(42), Cell.FromNumber(3.14), Cell.FromNumber(123456789012L), Cell.FromNumber(0), Cell.FromNumber(-99) },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Equal(1, read.Rows.Count);
            Assert.Equal(CellType.Number, read.Rows[0][0].Type);
            Assert.Equal(42, read.Rows[0][0].Number);
            Assert.Equal(3.14, read.Rows[0][1].Number, 0.001);
            Assert.Equal(123456789012L, (long)read.Rows[0][2].Number);
            Assert.Equal(0, read.Rows[0][3].Number);
            Assert.Equal(-99, read.Rows[0][4].Number);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Dates_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var date1 = new DateTime(2024, 1, 15);
            var date2 = new DateTime(1999, 12, 31, 23, 59, 59);

            var sheet = new SheetData
            {
                Headers = new() { "日期1", "日期2" },
                Rows = new()
                {
                    new Cell[] { Cell.FromDate(date1), Cell.FromDate(date2) },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Equal(1, read.Rows.Count);
            Assert.Equal(CellType.Date, read.Rows[0][0].Type);
            Assert.Equal(date1, read.Rows[0][0].Date);
            Assert.Equal(CellType.Date, read.Rows[0][1].Type);
            Assert.Equal(date2, read.Rows[0][1].Date);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Booleans_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                Headers = new() { "T", "F" },
                Rows = new()
                {
                    new Cell[] { Cell.FromBoolean(true), Cell.FromBoolean(false) },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Equal(1, read.Rows.Count);
            Assert.Equal(CellType.Boolean, read.Rows[0][0].Type);
            Assert.True(read.Rows[0][0].Boolean);
            Assert.Equal(CellType.Boolean, read.Rows[0][1].Type);
            Assert.False(read.Rows[0][1].Boolean);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void EmptyCells_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                Headers = new() { "A", "B", "C" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("x"), Cell.Empty, Cell.FromText("z") },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Equal(1, read.Rows.Count);
            Assert.Equal("x", read.Rows[0][0].Text);
            Assert.Equal(CellType.Empty, read.Rows[0][1].Type);
            Assert.Equal("z", read.Rows[0][2].Text);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void MixedTypes_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                Headers = new() { "文本", "数字", "日期", "布尔", "空" },
                Rows = new()
                {
                    new Cell[]
                    {
                        Cell.FromText("混合"),
                        Cell.FromNumber(123.45),
                        Cell.FromDate(new DateTime(2024, 6, 1)),
                        Cell.FromBoolean(true),
                        Cell.Empty,
                    },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Equal(1, read.Rows.Count);
            Assert.Equal(CellType.Text, read.Rows[0][0].Type);
            Assert.Equal("混合", read.Rows[0][0].Text);
            Assert.Equal(CellType.Number, read.Rows[0][1].Type);
            Assert.Equal(123.45, read.Rows[0][1].Number);
            Assert.Equal(CellType.Date, read.Rows[0][2].Type);
            Assert.Equal(new DateTime(2024, 6, 1), read.Rows[0][2].Date);
            Assert.Equal(CellType.Boolean, read.Rows[0][3].Type);
            Assert.True(read.Rows[0][3].Boolean);
            Assert.Equal(CellType.Empty, read.Rows[0][4].Type);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void NoHeaders_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "NoHeader",
                Rows = new()
                {
                    new Cell[] { Cell.FromText("A1"), Cell.FromText("B1") },
                    new Cell[] { Cell.FromText("A2"), Cell.FromText("B2") },
                },
            };
            XlsxWriter.Write(file, sheet);

            // 不把第一行当表头
            var read = XlsxReader.Read(file, 0, firstRowIsHeader: false);
            Assert.Empty(read.Headers);
            Assert.Equal(2, read.Rows.Count);
            Assert.Equal("A1", read.Rows[0][0].Text);
            Assert.Equal("B2", read.Rows[1][1].Text);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
