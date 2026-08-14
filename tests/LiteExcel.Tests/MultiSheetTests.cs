using LiteExcel;

namespace LiteExcel.Tests;

public class MultiSheetTests
{
    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"litexlsx_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void MultiSheet_WriteAndReadAll()
    {
        var file = GetTempFile();
        try
        {
            var sheets = new List<SheetData>
            {
                new()
                {
                    SheetName = "第一张",
                    Headers = new() { "A", "B" },
                    Rows = new() { new Cell[] { Cell.FromNumber(1), Cell.FromNumber(2) } },
                },
                new()
                {
                    SheetName = "第二张",
                    Headers = new() { "C", "D" },
                    Rows = new() { new Cell[] { Cell.FromText("x"), Cell.FromText("y") } },
                },
                new()
                {
                    SheetName = "Third",
                    Headers = new() { "E" },
                    Rows = new() { new Cell[] { Cell.FromBoolean(true) } },
                },
            };
            XlsxWriter.Write(file, sheets);

            // 列出 sheet 名
            var names = XlsxReader.GetSheetNames(file);
            Assert.Equal(3, names.Count);
            Assert.Equal("第一张", names[0]);
            Assert.Equal("第二张", names[1]);
            Assert.Equal("Third", names[2]);

            // 读全部
            var all = XlsxReader.ReadAll(file);
            Assert.Equal(3, all.Count);

            Assert.Equal("第一张", all[0].SheetName);
            Assert.Equal(1, all[0].Rows[0][0].Number);
            Assert.Equal(2, all[0].Rows[0][1].Number);

            Assert.Equal("第二张", all[1].SheetName);
            Assert.Equal("x", all[1].Rows[0][0].Text);
            Assert.Equal("y", all[1].Rows[0][1].Text);

            Assert.Equal("Third", all[2].SheetName);
            Assert.True(all[2].Rows[0][0].Boolean);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void ReadBySheetName()
    {
        var file = GetTempFile();
        try
        {
            var sheets = new List<SheetData>
            {
                new() { SheetName = "Alpha", Headers = new() { "X" }, Rows = new() { new Cell[] { Cell.FromNumber(10) } } },
                new() { SheetName = "Beta", Headers = new() { "Y" }, Rows = new() { new Cell[] { Cell.FromNumber(20) } } },
            };
            XlsxWriter.Write(file, sheets);

            var byName = XlsxReader.Read(file, "Beta");
            Assert.Equal("Beta", byName.SheetName);
            Assert.Equal(20, byName.Rows[0][0].Number);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void ReadByInvalidSheetName_Throws()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData { SheetName = "Only", Headers = new() { "A" }, Rows = new() };
            XlsxWriter.Write(file, sheet);

            Assert.Throws<LiteExcelException>(() => XlsxReader.Read(file, "Nonexistent"));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void ReadByInvalidIndex_Throws()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData { SheetName = "Only", Headers = new() { "A" }, Rows = new() };
            XlsxWriter.Write(file, sheet);

            Assert.Throws<ArgumentOutOfRangeException>(() => XlsxReader.Read(file, 5));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
