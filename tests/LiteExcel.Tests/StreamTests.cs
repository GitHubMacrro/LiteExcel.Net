using System.Data;
using System.IO;
using LiteExcel;

namespace LiteExcel.Tests;

public class StreamTests
{
    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"litexlsx_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void MemoryStream_RoundTrip_TextAndNumber()
    {
        var sheet = new SheetData
        {
            SheetName = "Test",
            Headers = new() { "A", "B" },
            Rows = new()
            {
                new Cell[] { Cell.FromText("hello"), Cell.FromNumber(42) },
                new Cell[] { Cell.FromText("中文"), Cell.FromNumber(3.14) },
            },
        };

        using var ms = new MemoryStream();
        XlsxWriter.Write(ms, sheet);
        ms.Position = 0;

        var read = XlsxReader.Read(ms, 0);
        Assert.Equal("Test", read.SheetName);
        Assert.Equal(2, read.Headers.Count);
        Assert.Equal("A", read.Headers[0]);
        Assert.Equal("B", read.Headers[1]);

        Assert.Equal(2, read.Rows.Count);
        Assert.Equal("hello", read.Rows[0][0].Text);
        Assert.Equal(42, read.Rows[0][1].Number);
        Assert.Equal("中文", read.Rows[1][0].Text);
        Assert.Equal(3.14, read.Rows[1][1].Number, 0.001);
    }

    [Fact]
    public void MemoryStream_MultiSheet_RoundTrip()
    {
        var sheets = new List<SheetData>
        {
            new() { SheetName = "Alpha", Headers = new() { "X" }, Rows = new() { new Cell[] { Cell.FromNumber(1) } } },
            new() { SheetName = "Beta", Headers = new() { "Y" }, Rows = new() { new Cell[] { Cell.FromText("y") } } },
        };

        using var ms = new MemoryStream();
        XlsxWriter.Write(ms, sheets);
        ms.Position = 0;

        var all = XlsxReader.ReadAll(ms);
        Assert.Equal(2, all.Count);
        Assert.Equal("Alpha", all[0].SheetName);
        Assert.Equal(1, all[0].Rows[0][0].Number);
        Assert.Equal("Beta", all[1].SheetName);
        Assert.Equal("y", all[1].Rows[0][0].Text);
    }

    [Fact]
    public void ReadWithProgress_CallbacksFrom1ToTotal()
    {
        var file = GetTempFile();
        try
        {
            const int rowCount = 100;
            var rows = new List<IReadOnlyList<Cell>>(rowCount);
            for (int i = 0; i < rowCount; i++)
            {
                rows.Add(new Cell[] { Cell.FromNumber(i), Cell.FromText($"Row{i}") });
            }
            var sheet = new SheetData
            {
                SheetName = "Data",
                Headers = new() { "ID", "Name" },
                Rows = rows,
            };
            XlsxWriter.Write(file, sheet);

            var currents = new List<int>();
            int total = -1;
            XlsxReader.ReadWithProgress(file, 0, (current, t) =>
            {
                currents.Add(current);
                total = t;
            });

            Assert.Equal(rowCount, total);
            Assert.Equal(rowCount, currents.Count);
            // current 从 1 递增到 total
            for (int i = 0; i < currents.Count; i++)
            {
                Assert.Equal(i + 1, currents[i]);
            }
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void StreamWrite_PathRead_Consistent()
    {
        var sheet = new SheetData
        {
            SheetName = "S",
            Headers = new() { "X", "Y" },
            Rows = new() { new Cell[] { Cell.FromNumber(1), Cell.FromText("a") } },
        };

        var file = GetTempFile();
        try
        {
            using (var fs = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                XlsxWriter.Write(fs, sheet);
            }

            var read = XlsxReader.Read(file, 0);
            Assert.Equal("S", read.SheetName);
            Assert.Equal("X", read.Headers[0]);
            Assert.Equal(1, read.Rows[0][0].Number);
            Assert.Equal("a", read.Rows[0][1].Text);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void PathWrite_StreamRead_Consistent()
    {
        var sheet = new SheetData
        {
            SheetName = "PS",
            Headers = new() { "C1", "C2" },
            Rows = new() { new Cell[] { Cell.FromText("val"), Cell.FromNumber(99) } },
        };

        var file = GetTempFile();
        try
        {
            XlsxWriter.Write(file, sheet);

            using var fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var read = XlsxReader.Read(fs, 0);
            Assert.Equal("PS", read.SheetName);
            Assert.Equal("val", read.Rows[0][0].Text);
            Assert.Equal(99, read.Rows[0][1].Number);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void GetSheetNames_Stream()
    {
        var sheets = new List<SheetData>
        {
            new() { SheetName = "First", Headers = new() { "A" }, Rows = new() { new Cell[] { Cell.FromNumber(1) } } },
            new() { SheetName = "Second", Headers = new() { "B" }, Rows = new() { new Cell[] { Cell.FromNumber(2) } } },
            new() { SheetName = "Third", Headers = new() { "C" }, Rows = new() { new Cell[] { Cell.FromNumber(3) } } },
        };

        using var ms = new MemoryStream();
        XlsxWriter.Write(ms, sheets);
        ms.Position = 0;

        var names = XlsxReader.GetSheetNames(ms);
        Assert.Equal(3, names.Count);
        Assert.Equal("First", names[0]);
        Assert.Equal("Second", names[1]);
        Assert.Equal("Third", names[2]);
    }

    [Fact]
    public void MemoryStream_1000Rows_LargeFile()
    {
        const int rowCount = 1000;
        var rows = new List<IReadOnlyList<Cell>>(rowCount);
        for (int i = 0; i < rowCount; i++)
        {
            rows.Add(new Cell[] { Cell.FromNumber(i), Cell.FromText($"Item{i:D4}") });
        }
        var sheet = new SheetData
        {
            SheetName = "Big",
            Headers = new() { "ID", "Name" },
            Rows = rows,
        };

        using var ms = new MemoryStream();
        XlsxWriter.Write(ms, sheet);
        ms.Position = 0;

        var read = XlsxReader.Read(ms, 0);
        Assert.Equal(rowCount, read.Rows.Count);
        Assert.Equal(0, read.Rows[0][0].Number);
        Assert.Equal("Item0000", read.Rows[0][1].Text);
        Assert.Equal(999, read.Rows[999][0].Number);
        Assert.Equal("Item0999", read.Rows[999][1].Text);
    }

    [Fact]
    public void StreamRows_Stream()
    {
        var sheet = new SheetData
        {
            SheetName = "Stream",
            Headers = new() { "ID", "Val" },
            Rows = new()
            {
                new Cell[] { Cell.FromNumber(10), Cell.FromText("a") },
                new Cell[] { Cell.FromNumber(20), Cell.FromText("b") },
                new Cell[] { Cell.FromNumber(30), Cell.FromText("c") },
            },
        };

        using var ms = new MemoryStream();
        XlsxWriter.Write(ms, sheet);
        ms.Position = 0;

        var rows = new List<IReadOnlyList<Cell>>();
        XlsxReader.StreamRows(ms, "Stream", row => rows.Add(row));

        Assert.Equal(3, rows.Count);
        Assert.Equal(10, rows[0][0].Number);
        Assert.Equal("a", rows[0][1].Text);
        Assert.Equal(30, rows[2][0].Number);
        Assert.Equal("c", rows[2][1].Text);
    }

    [Fact]
    public void ReadAsDataTable_Stream()
    {
        var sheet = new SheetData
        {
            SheetName = "DT",
            Headers = new() { "Name", "Age" },
            Rows = new()
            {
                new Cell[] { Cell.FromText("Alice"), Cell.FromNumber(30) },
                new Cell[] { Cell.FromText("Bob"), Cell.FromNumber(25) },
            },
        };

        using var ms = new MemoryStream();
        XlsxWriter.Write(ms, sheet);
        ms.Position = 0;

        var dt = XlsxReader.ReadAsDataTable(ms, 0);
        Assert.Equal("DT", dt.TableName);
        Assert.Equal(2, dt.Columns.Count);
        Assert.Equal("Name", dt.Columns[0].ColumnName);
        Assert.Equal("Age", dt.Columns[1].ColumnName);
        Assert.Equal(2, dt.Rows.Count);
        Assert.Equal("Alice", dt.Rows[0][0]);
        Assert.Equal(30.0, dt.Rows[0][1]);
        Assert.Equal("Bob", dt.Rows[1][0]);
        Assert.Equal(25.0, dt.Rows[1][1]);
    }

    [Fact]
    public void ReadBySheetName_Stream()
    {
        var sheets = new List<SheetData>
        {
            new() { SheetName = "Alpha", Headers = new() { "X" }, Rows = new() { new Cell[] { Cell.FromNumber(10) } } },
            new() { SheetName = "Beta", Headers = new() { "Y" }, Rows = new() { new Cell[] { Cell.FromNumber(20) } } },
        };

        using var ms = new MemoryStream();
        XlsxWriter.Write(ms, sheets);
        ms.Position = 0;

        var byName = XlsxReader.Read(ms, "Beta");
        Assert.Equal("Beta", byName.SheetName);
        Assert.Equal(20, byName.Rows[0][0].Number);
    }

    [Fact]
    public void ReadInvalidIndex_Stream_Throws()
    {
        var sheet = new SheetData { SheetName = "Only", Headers = new() { "A" }, Rows = new() };
        using var ms = new MemoryStream();
        XlsxWriter.Write(ms, sheet);
        ms.Position = 0;

        Assert.Throws<ArgumentOutOfRangeException>(() => XlsxReader.Read(ms, 5));
    }
}