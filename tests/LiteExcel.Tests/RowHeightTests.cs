using LiteExcel;

namespace LiteExcel.Tests;

public class RowHeightTests
{
    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"litexlsx_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void RowHeight_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "Test",
                Headers = new() { "A", "B" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("x"), Cell.FromText("y") },
                    new Cell[] { Cell.FromText("z"), Cell.FromText("w") },
                },
                RowHeights = new()
                {
                    { 0, 20.5 },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.NotNull(read.RowHeights);
            Assert.True(read.RowHeights!.ContainsKey(0));
            Assert.Equal(20.5, read.RowHeights[0], 0.001);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void MultipleRowHeights_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "Test",
                Headers = new() { "A" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("r0") },
                    new Cell[] { Cell.FromText("r1") },
                    new Cell[] { Cell.FromText("r2") },
                },
                RowHeights = new()
                {
                    { 0, 15.0 },
                    { 1, 30.25 },
                    { 2, 45.75 },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.NotNull(read.RowHeights);
            Assert.Equal(3, read.RowHeights!.Count);
            Assert.Equal(15.0, read.RowHeights[0], 0.001);
            Assert.Equal(30.25, read.RowHeights[1], 0.001);
            Assert.Equal(45.75, read.RowHeights[2], 0.001);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void NoRowHeights_DoesNotOutputHtAttribute()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "Test",
                Headers = new() { "A" },
                Rows = new() { new Cell[] { Cell.FromText("x") } },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Null(read.RowHeights);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void PartialRowHeights_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "Test",
                Headers = new() { "A" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("r0") },
                    new Cell[] { Cell.FromText("r1") },
                    new Cell[] { Cell.FromText("r2") },
                    new Cell[] { Cell.FromText("r3") },
                },
                RowHeights = new()
                {
                    // 只有第 0 和第 2 行有行高，其它行没有
                    { 0, 18.0 },
                    { 2, 36.0 },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.NotNull(read.RowHeights);
            Assert.Equal(2, read.RowHeights!.Count);
            Assert.True(read.RowHeights.ContainsKey(0));
            Assert.True(read.RowHeights.ContainsKey(2));
            Assert.False(read.RowHeights.ContainsKey(1));
            Assert.False(read.RowHeights.ContainsKey(3));
            Assert.Equal(18.0, read.RowHeights[0], 0.001);
            Assert.Equal(36.0, read.RowHeights[2], 0.001);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void RowHeights_NoHeaders()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "Test",
                Rows = new()
                {
                    new Cell[] { Cell.FromText("r0") },
                    new Cell[] { Cell.FromText("r1") },
                },
                RowHeights = new()
                {
                    { 1, 25.0 },
                },
            };
            XlsxWriter.Write(file, sheet);

            // 不把第一行当表头，所以 RowHeights 索引直接对应 Rows
            var read = XlsxReader.Read(file, 0, firstRowIsHeader: false);
            Assert.NotNull(read.RowHeights);
            Assert.True(read.RowHeights!.ContainsKey(1));
            Assert.Equal(25.0, read.RowHeights[1], 0.001);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void RowHeights_IntegerValue()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "Test",
                Headers = new() { "A" },
                Rows = new() { new Cell[] { Cell.FromText("x") } },
                RowHeights = new()
                {
                    { 0, 30 },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.NotNull(read.RowHeights);
            Assert.Equal(30.0, read.RowHeights![0], 0.001);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void RowHeights_MultiSheet_Independent()
    {
        var file = GetTempFile();
        try
        {
            var sheets = new List<SheetData>
            {
                new()
                {
                    SheetName = "Sheet1",
                    Headers = new() { "A" },
                    Rows = new() { new Cell[] { Cell.FromText("x") } },
                    RowHeights = new() { { 0, 12.5 } },
                },
                new()
                {
                    SheetName = "Sheet2",
                    Headers = new() { "B" },
                    Rows = new() { new Cell[] { Cell.FromText("y") } },
                    RowHeights = new() { { 0, 40.0 } },
                },
            };
            XlsxWriter.Write(file, sheets);

            var all = XlsxReader.ReadAll(file);
            Assert.Equal(2, all.Count);
            Assert.NotNull(all[0].RowHeights);
            Assert.Equal(12.5, all[0].RowHeights![0], 0.001);
            Assert.NotNull(all[1].RowHeights);
            Assert.Equal(40.0, all[1].RowHeights![0], 0.001);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}