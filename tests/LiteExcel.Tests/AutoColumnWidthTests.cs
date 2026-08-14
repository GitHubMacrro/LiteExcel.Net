using LiteExcel;

namespace LiteExcel.Tests;

public class AutoColumnWidthTests
{
    [Fact]
    public void AutoColumnWidths_EstimatesCorrectly()
    {
        var sheet = new SheetData
        {
            Headers = new() { "Name", "Age" },
            Rows = new()
            {
                new Cell[] { Cell.FromText("Alice"), Cell.FromNumber(30) },
                new Cell[] { Cell.FromText("Bob"), Cell.FromNumber(100) },
            },
        };

        XlsxWriter.AutoColumnWidths(sheet);

        Assert.NotNull(sheet.ColumnWidths);
        Assert.Equal(2, sheet.ColumnWidths!.Count);
        // "Alice" = 5 chars, "Name" = 4 chars -> max 5, but min 8
        Assert.Equal(8, sheet.ColumnWidths[0]);
        // "100" = 3, "Age" = 3 -> max 3, but min 8
        Assert.Equal(8, sheet.ColumnWidths[1]);
    }

    [Fact]
    public void AutoColumnWidths_ChineseCharsCountAs2()
    {
        var sheet = new SheetData
        {
            Headers = new() { "中文名称" },
            Rows = new()
            {
                new Cell[] { Cell.FromText("你好世界测试") },
            },
        };

        XlsxWriter.AutoColumnWidths(sheet);

        Assert.NotNull(sheet.ColumnWidths);
        // "中文名称" = 4*2 = 8, "你好世界测试" = 6*2 = 12 -> max 12
        Assert.Equal(12, sheet.ColumnWidths![0]);
    }

    [Fact]
    public void AutoColumnWidths_CappedAt50()
    {
        var longText = new string('A', 60);
        var sheet = new SheetData
        {
            Headers = new() { "col" },
            Rows = new() { new Cell[] { Cell.FromText(longText) } },
        };

        XlsxWriter.AutoColumnWidths(sheet);

        Assert.NotNull(sheet.ColumnWidths);
        Assert.Equal(50, sheet.ColumnWidths![0]);
    }

    [Fact]
    public void AutoColumnWidths_RoundTrip()
    {
        var file = Path.Combine(Path.GetTempPath(), $"litexlsx_{Guid.NewGuid():N}.xlsx");
        try
        {
            var sheet = new SheetData
            {
                SheetName = "Test",
                Headers = new() { "产品名称", "数量" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("苹果手机"), Cell.FromNumber(100) },
                    new Cell[] { Cell.FromText("笔记本电脑"), Cell.FromNumber(50) },
                },
            };
            XlsxWriter.AutoColumnWidths(sheet);
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Equal(2, read.Rows.Count);
            Assert.Equal("苹果手机", read.Rows[0][0].Text);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void AutoColumnWidths_EmptySheet_NoCrash()
    {
        var sheet = new SheetData
        {
            SheetName = "Test",
            Headers = new(),
            Rows = new(),
        };

        XlsxWriter.AutoColumnWidths(sheet);
        // No assertion needed - just shouldn't crash
    }
}